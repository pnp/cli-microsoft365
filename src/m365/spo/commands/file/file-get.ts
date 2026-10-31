import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  url: z.string().optional(),
  id: z.uuid().optional().alias('i'),
  asString: z.boolean().optional(),
  asListItem: z.boolean().optional(),
  asFile: z.boolean().optional(),
  path: z.string()
    .refine(p => fs.existsSync(path.dirname(p)), {
      error: 'Specified path where to save the file does not exits'
    })
    .optional()
    .alias('p'),
  withPermissions: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified file';
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
      })
      .refine(opts => !opts.asFile || !!opts.path, {
        error: 'The path should be specified when the --asFile option is used',
        path: ['path'],
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => [opts.asFile, opts.asListItem, opts.asString].filter(o => o === true).length <= 1, {
        error: 'Specify to retrieve the file either as file, list item or string but not multiple'
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving file from site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';
    let options: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.url) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)`;
    }

    if (args.options.asListItem) {
      options = '?$expand=ListItemAllFields';
    }
    else if (args.options.asFile || args.options.asString) {
      options = '/$value';
    }

    if (args.options.url) {
      if (options.indexOf('?') < 0) {
        options += '?';
      }
      else {
        options += '&';
      }

      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      options += `@f='${formatting.encodeQueryParameter(serverRelativePath)}'`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl + options,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      // Use stream for --asFile, so binary data is piped to disk as-is and
      // isn't encoded to utf8 (which would corrupt it). Use text for
      // --asString. Use json for all other cases.
      responseType: args.options.asFile ? 'stream' : args.options.asString ? 'text' : 'json'
    };

    try {
      const file = await request.get<any>(requestOptions);

      if (args.options.asFile && args.options.path) {
        // Not possible to use async/await for this promise
        await new Promise<void>((resolve, reject) => {
          const writer = fs.createWriteStream(args.options.path as string);
          file.data.pipe(writer);

          writer.on('error', err => {
            reject(err);
          });
          writer.on('close', async () => {
            const filePath = args.options.path as string;
            if (this.verbose) {
              await logger.logToStderr(`File saved to path ${filePath}`);
            }
            return resolve();
          });
        });
      }
      else {
        if (args.options.asString) {
          await logger.log(file.toString());
        }
        else {
          const fileProperties: FileProperties = JSON.parse(JSON.stringify(file));

          if (args.options.withPermissions) {
            requestOptions.url = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${file.ServerRelativeUrl}')/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
            const response = await request.get<{ value: any[] }>(requestOptions);
            response.value.forEach((r: any) => {
              r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
            });
            fileProperties.RoleAssignments = response.value;
            if (args.options.asListItem) {
              fileProperties.ListItemAllFields.RoleAssignments = response.value;
            }
          }

          if (args.options.asListItem) {
            delete fileProperties.ListItemAllFields.ID;
          }

          await logger.log(args.options.asListItem ? fileProperties.ListItemAllFields : fileProperties);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileGetCommand();