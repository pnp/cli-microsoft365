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

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileCheckoutUndoCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_CHECKOUT_UNDO;
  }

  public get description(): string {
    return 'Discards a checked out file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileId, opts.fileUrl].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileId' or 'fileUrl', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileId', 'fileUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const undoCheckout = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Undoing checkout for file ${args.options.fileId || args.options.fileUrl} on web ${args.options.webUrl}`);
        }

        let requestUrl: string = `${args.options.webUrl}/_api/web/`;

        if (args.options.fileId) {
          requestUrl += `getFileById('${args.options.fileId}')`;
        }
        else if (args.options.fileUrl) {
          const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
          requestUrl += `getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
        }

        requestUrl += '/undocheckout';

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await undoCheckout();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to undo the checkout for file ${args.options.fileId || args.options.fileUrl}?` });

      if (result) {
        await undoCheckout();
      }
    }
  }
}

export default new SpoFileCheckoutUndoCommand();