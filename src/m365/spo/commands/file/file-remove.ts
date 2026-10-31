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
  id: z.uuid().optional().alias('i'),
  url: z.string().optional(),
  permanent: z.boolean().optional(),
  bypassSharedLock: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified file';
  }

  public alias(): string[] | undefined {
    return [commands.PAGE_TEMPLATE_REMOVE];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.url].filter(o => o !== undefined).length === 1, {
        error: `Specify 'id' or 'url', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'url']
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeFile = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing file in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string;

      if (args.options.id) {
        requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${formatting.encodeQueryParameter(args.options.id as string)}')`;
      }
      else {
        const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url!);
        requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
      }

      if (!args.options.permanent) {
        requestUrl += `/recycle()`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.bypassSharedLock) {
        requestOptions.headers!.Prefer = 'bypass-shared-lock';
      }

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeFile();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.permanent ? 'permanently remove' : 'recycle'} the file ${args.options.id || args.options.url} located in site ${args.options.webUrl}?` });

      if (result) {
        await removeFile();
      }
    }
  }
}

export default new SpoFileRemoveCommand();