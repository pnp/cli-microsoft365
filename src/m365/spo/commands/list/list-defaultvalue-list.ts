import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { DOMParser } from '@xmldom/xmldom';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

interface DefaultColumnValue {
  fieldName: string;
  fieldValue: string;
  folderUrl: string
}

const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    listId: zod.alias('i', z.string().optional()
      .refine(id => id === undefined || validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      }))
    ),
    listTitle: zod.alias('t', z.string().optional()),
    listUrl: z.string().optional(),
    folderUrl: z.string().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListDefaultValueListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_DEFAULTVALUE_LIST;
  }

  public get description(): string {
    return 'Retrieves default column values for a specific document library';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: listId, listTitle, listUrl.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving default column values for list '${args.options.listId || args.options.listTitle || args.options.listUrl}'...`);
        await logger.logToStderr('Retrieving list information...');
      }

      const listServerRelUrl = await this.getServerRelativeListUrl(args.options);

      if (this.verbose) {
        await logger.logToStderr('Retrieving default column values...');
      }

      let defaultValues: DefaultColumnValue[];
      try {
        const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
        defaultValues = this.convertXmlToJson(defaultValuesXml);
      }
      catch (err: any) {
        if (err.status !== 404) {
          throw err;
        }
        // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
        defaultValues = [];
      }

      if (args.options.folderUrl) {
        const serverRelFolderUrl = urlUtil.removeTrailingSlashes(urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl));
        defaultValues = defaultValues.filter(d => d.folderUrl.toLowerCase() === serverRelFolderUrl.toLowerCase());
      }
      await logger.log(defaultValues);
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
      responseType: 'json'
    };
    const defaultValuesXml = await request.get<string>(requestOptions);
    return defaultValuesXml;
  }

  private convertXmlToJson(xml: string): DefaultColumnValue[] {
    const results: DefaultColumnValue[] = [];

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    const folderLinks = doc.getElementsByTagName('a');
    for (let i = 0; i < folderLinks.length; i++) {
      const folderUrl = folderLinks[i].getAttribute('href')!;
      const defaultValues = folderLinks[i].getElementsByTagName('DefaultValue');

      for (let j = 0; j < defaultValues.length; j++) {
        const fieldName = defaultValues[j].getAttribute('FieldName')!;
        const fieldValue = defaultValues[j].textContent!;

        results.push({
          fieldName: fieldName,
          fieldValue: fieldValue,
          folderUrl: decodeURIComponent(folderUrl)
        });
      }
    }

    return results;
  }
}

export default new SpoListDefaultValueListCommand();