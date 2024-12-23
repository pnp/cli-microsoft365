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
import { cli } from '../../../../cli/cli.js';

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
    fieldName: z.string().optional(),
    folderUrl: z.string().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListDefaultValueClearCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_DEFAULTVALUE_CLEAR;
  }

  public get description(): string {
    return 'Clears default column values for a specific document library';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: listId, listTitle, listUrl.'
      })
      .refine(options => (options.fieldName !== undefined) !== (options.folderUrl !== undefined) || (options.fieldName === undefined && options.folderUrl === undefined), {
        message: `Specify 'fieldName' or 'folderUrl', but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear all default values${args.options.fieldName ? ` for field '${args.options.fieldName}'` : args.options.folderUrl ? ` for folder ${args.options.folderUrl}` : ''}?` });
      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Clearing all default column values${args.options.fieldName ? ` for field ${args.options.fieldName}` : args.options.folderUrl ? `for folder '${args.options.folderUrl}'` : ''}...`);
        await logger.logToStderr(`Getting server-relative URL of the list...`);
      }

      const listServerRelUrl = await this.getServerRelativeListUrl(args.options);

      if (this.verbose) {
        await logger.logToStderr(`List server-relative URL: ${listServerRelUrl}`);
        await logger.logToStderr(`Getting default column values...`);
      }

      const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);

      if (defaultValuesXml === null) {
        if (this.verbose) {
          await logger.logToStderr(`No default column values found.`);
        }
        return;
      }

      const trimmedXml = this.removeFieldsFromXml(defaultValuesXml, args.options);
      await this.uploadDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl, trimmedXml);
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
      requestOptions.url += `/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }
    else if (options.listId) {
      requestOptions.url += `/Lists('${options.listId}')`;
    }
    else if (options.listTitle) {
      requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }

    requestOptions.url += '?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate';

    try {
      const response = await request.get<{ BaseTemplate: number; RootFolder: { ServerRelativeUrl: string } }>(requestOptions);

      if (response.BaseTemplate !== 101) {
        throw `The specified list is not a document library.`;
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

  private async getDefaultColumnValuesXml(webUrl: string, listServerRelUrl: string): Promise<string | null> {
    try {
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
    catch (err: any) {
      // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
      if (err.status === 404) {
        return null;
      }

      throw err;
    }
  }

  private removeFieldsFromXml(xml: string, options: Options): string {
    if (!options.fieldName && !options.folderUrl) {
      return '<MetadataDefaults />';
    }

    let folderUrlToRemove = null;

    if (options.folderUrl) {
      folderUrlToRemove = urlUtil.removeTrailingSlashes(urlUtil.getServerRelativePath(options.webUrl, options.folderUrl));
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    const folderLinks = doc.getElementsByTagName('a');

    for (let i = 0; i < folderLinks.length; i++) {
      const folderNode = folderLinks[i];
      const folderUrl = folderNode.getAttribute('href')!;

      if (folderUrlToRemove && folderUrlToRemove.toLowerCase() === decodeURIComponent(folderUrl).toLowerCase()) {
        folderNode.parentNode!.removeChild(folderNode);
        break;
      }
      else if (options.fieldName) {
        const defaultValues = folderNode.getElementsByTagName('DefaultValue');

        for (let j = 0; j < defaultValues.length; j++) {
          const defaultValueNode = defaultValues[j];
          const fieldName = defaultValueNode.getAttribute('FieldName')!;

          if (fieldName.toLowerCase() === options.fieldName!.toLowerCase()) {
            // Remove the entire folder node if it becomes empty
            if (folderNode.childNodes.length === 1) {
              folderNode.parentNode!.removeChild(defaultValueNode.parentNode!);
            }
            else {
              folderNode.removeChild(defaultValueNode);
            }
            break;
          }
        }
      }
    }

    return doc.toString();
  }

  private async uploadDefaultColumnValuesXml(webUrl: string, listServerRelUrl: string, xml: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'If-Match': '*'
      },
      responseType: 'json',
      data: xml
    };

    await request.put(requestOptions);
  }
}

export default new SpoListDefaultValueClearCommand();