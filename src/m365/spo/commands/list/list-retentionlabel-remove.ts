import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { ListInstance } from './ListInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listTitle: z.string().alias('t').optional(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i')
    .optional(),
  listUrl: z.string().alias('l').optional(),
  force: z.boolean().alias('f').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListRetentionLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clears the retention label on the specified list or library.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeListRetentionLabel(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });

      if (result) {
        await this.removeListRetentionLabel(logger, args);
      }
    }
  }

  private async removeListRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Clears the retention label from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }

    try {
      const listServerRelativeUrl: string = await this.getListServerRelativeUrl(args, logger);
      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);

      await spo.removeDefaultRetentionLabelFromList(args.options.webUrl, listAbsoluteUrl, logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListServerRelativeUrl(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr('Getting the list server relative URL');
    }

    if (args.options.listUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.listId) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listInstance = await request.get<ListInstance>(requestOptions);
    return listInstance.RootFolder.ServerRelativeUrl;
  }
}

export default new SpoListRetentionLabelRemoveCommand();