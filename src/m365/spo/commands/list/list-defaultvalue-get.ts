import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { spo } from '../../../../utils/spo.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { DefaultColumnValue } from "./DefaultColumnValue";

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.uuid().optional().alias('i'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  fieldName: z.string(),
  folderUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListDefaultValueGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_DEFAULTVALUE_GET;
  }

  public get description(): string {
    return 'Gets a specific default column value from a list';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: listId, listTitle, listUrl.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving default column value for field '${args.options.fieldName}' in list '${args.options.listId || args.options.listTitle || args.options.listUrl}'...`);
        await logger.logToStderr('Retrieving list information...');
      }

      const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
      if (this.verbose) {
        await logger.logToStderr('Retrieving default column values...');
      }

      let defaultValues: DefaultColumnValue[];
      try {
        const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
        defaultValues = spo.convertDefaultColumnXmlToJson(defaultValuesXml);
      }
      catch (err: any) {
        if (err.status !== 404) {
          throw err;
        }
        // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
        defaultValues = [];
      }
      defaultValues = defaultValues.filter(d => d.fieldName.toLowerCase() === args.options.fieldName.toLowerCase());

      if (args.options.folderUrl) {
        const serverRelFolderUrl = urlUtil.removeTrailingSlashes(urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl));
        defaultValues = defaultValues.filter(d => d.folderUrl.toLowerCase() === serverRelFolderUrl.toLowerCase());
      }
      else {
        defaultValues = defaultValues.filter(d => d.folderUrl.toLowerCase() === listServerRelUrl.toLowerCase());
      }

      if (defaultValues.length === 0) {
        throw `No default column value found for field '${args.options.fieldName}'${args.options.folderUrl ? ` in folder '${args.options.folderUrl}'` : ''}.`;
      }

      await logger.log(defaultValues[0]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getServerRelativeListUrl(options: Options): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/Web`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (options.listUrl) {
      const serverRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestOptions.url += `/GetList('${serverRelativeUrl}')`;
    }
    else if (options.listId) {
      requestOptions.url += `/Lists('${options.listId}')`;
    }
    else if (options.listTitle) {
      requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }

    requestOptions.url += '?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate';

    try {
      const response = await request.get<{ RootFolder: { ServerRelativeUrl: string }, BaseTemplate: number }>(requestOptions);
      if (response.BaseTemplate !== 101) {
        throw `List '${options.listId || options.listTitle || options.listUrl}' is not a document library.`;
      }
      return response.RootFolder.ServerRelativeUrl;
    }
    catch (error: any) {
      if (error.status === 404) {
        throw `List '${options.listId || options.listTitle || options.listUrl}' was not found.`;
      }

      throw error;
    }
  }

  private async getDefaultColumnValuesXml(webUrl: string, listServerRelUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'text'
    };
    const defaultValuesXml = await request.get<string>(requestOptions);
    return defaultValuesXml;
  }
}

export default new SpoListDefaultValueGetCommand();