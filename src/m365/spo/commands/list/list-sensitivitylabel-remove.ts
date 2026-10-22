import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: 'Specify a valid SharePoint site URL.'
    })
    .alias('u'),
  listTitle: z.string().optional().alias('t'),
  listId: z.uuid().optional().alias('l'),
  listUrl: z.string().optional(),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListSensitivityLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SENSITIVITYLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clears a default sensitivity label from a document library';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: listId, listTitle, or listUrl.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSensitivityLabel = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing the sensitivity label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' in site at ${args.options.webUrl}...`);
        }

        let requestUrl: string = `${args.options.webUrl}/_api/web`;

        if (args.options.listId) {
          requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
          requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
          const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
          requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata',
            'if-match': '*'
          },
          data: { 'DefaultSensitivityLabelForLibrary': '' },
          responseType: 'json'
        };

        await request.patch<any>(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeSensitivityLabel();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the sensitivity label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });

      if (result) {
        await removeSensitivityLabel();
      }
    }
  }
}

export default new SpoListSensitivityLabelRemoveCommand();
