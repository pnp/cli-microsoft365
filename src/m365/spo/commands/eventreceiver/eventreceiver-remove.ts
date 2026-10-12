import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import getCommand, { Options as SpoEventReceiverGetOptions } from './eventreceiver-get.js';
import { EventReceiver } from './EventReceiver.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID.`
  }).optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID.`
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  scope: z.enum(['web', 'site']).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoEventreceiverRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.EVENTRECEIVER_REMOVE;
  }

  public get description(): string {
    return 'Removes event receivers for the specified web, site, or list.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => {
        const listOptions = [opts.listId, opts.listTitle, opts.listUrl];
        return !(listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1);
      }, {
        message: `Specify either list id or title or list url`,
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => !(opts.scope === 'site' && (opts.listId || opts.listUrl || opts.listTitle)), {
        message: 'Scope cannot be set to site when retrieving list event receivers.',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeEventReceiver(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove event receiver with ${args.options.id ? `id ${args.options.id}` : `name ${args.options.name}`}?` });

      if (result) {
        await this.removeEventReceiver(args.options);
      }
    }
  }

  public async removeEventReceiver(options: Options): Promise<void> {
    try {
      let requestUrl = `${options.webUrl}/_api/${options.scope || 'web'}`;

      if (options.listId) {
        requestUrl += `/lists('${options.listId}')`;
      }
      else if (options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const rerId = await this.getEventReceiverId(options);
      requestUrl += `/eventreceivers('${rerId}')`;

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getEventReceiverId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const getOptions: SpoEventReceiverGetOptions = {
      webUrl: options.webUrl,
      listId: options.listId,
      listTitle: options.listTitle,
      listUrl: options.listUrl,
      scope: options.scope,
      id: options.id,
      name: options.name,
      debug: this.debug,
      verbose: this.verbose
    };

    const commandOutput = await cli.executeCommandWithOutput(getCommand as Command, { options: { ...getOptions, _: [] } });
    const eventReceiver: EventReceiver = JSON.parse(commandOutput.stdout);

    return eventReceiver.ReceiverId;
  }
}

export default new SpoEventreceiverRemoveCommand();