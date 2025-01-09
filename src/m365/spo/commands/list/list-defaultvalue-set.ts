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
    fieldName: z.string(),
    fieldValue: z.string()
      .refine(value => value !== '', `The value cannot be empty. Use 'spo list defaultvalue remove' to remove a default column value.`),
    folderUrl: z.string().optional()
      .refine(url => url === undefined || (!url.includes('#') && !url.includes('%')), 'Due to limitations in SharePoint Online, setting default column values for folders with a # or % character in their path is not supported.')
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListDefaultValueSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_DEFAULTVALUE_SET;
  }

  public get description(): string {
    return 'Sets default column values for a specific document library';
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
        await logger.logToStderr(`Setting default column value '${args.options.fieldValue}' for field '${args.options.fieldName}'...`);
        await logger.logToStderr(`Getting server-relative URL of the list...`);
      }

      const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
      let folderUrl = listServerRelUrl;

      if (args.options.folderUrl) {
        if (this.verbose) {
          await logger.logToStderr(`Getting server-relative URL of folder '${args.options.folderUrl}'...`);
        }

        // Casing of the folder URL is important, let's retrieve the correct URL
        const serverRelativeFolderUrl = urlUtil.getServerRelativePath(args.options.webUrl, urlUtil.removeTrailingSlashes(args.options.folderUrl));
        folderUrl = await this.getCorrectFolderUrl(args.options.webUrl, serverRelativeFolderUrl);
      }

      if (this.verbose) {
        await logger.logToStderr(`Getting default column values...`);
      }

      const defaultValuesXml = await this.ensureDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
      const modifiedXml = await this.updateFieldValueXml(logger, defaultValuesXml, args.options.fieldName, args.options.fieldValue, folderUrl);
      await this.uploadDefaultColumnValuesXml(logger, args.options.webUrl, listServerRelUrl, modifiedXml);
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

  private async getCorrectFolderUrl(webUrl: string, folderUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      // Using ListItemAllFields endpoint because GetFolderByServerRelativePath doesn't return the correctly cased URL
      url: `${webUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ FileRef: string }>(requestOptions);

    if (!response.FileRef) {
      throw `Folder '${folderUrl}' was not found.`;
    }

    return response.FileRef;
  }

  private async ensureDefaultColumnValuesXml(webUrl: string, listServerRelUrl: string): Promise<string> {
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
      if (err.status !== 404) {
        throw err;
      }

      // For lists that have never had default column values set, the client_LocationBasedDefaults.html file does not exist.
      // In this case, we need to create the file with blank default metadata.
      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms')}')/Files/Add(url='client_LocationBasedDefaults.html', overwrite=false)`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'text/plain'
        },
        responseType: 'json',
        data: '<MetadataDefaults />'
      };

      await request.post(requestOptions);
      return requestOptions.data;
    }
  }

  private async updateFieldValueXml(logger: Logger, xml: string, fieldName: string, fieldValue: string, folderUrl: string): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Modifying default column values...`);
    }
    // Encode all spaces in the folder URL
    const encodedFolderUrl = folderUrl.replace(/ /g, '%20');

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    // Create a new DefaultValue node
    const newDefaultValueNode = doc.createElement('DefaultValue');
    newDefaultValueNode.setAttribute('FieldName', fieldName);
    newDefaultValueNode.textContent = fieldValue;

    const folderLinks = doc.getElementsByTagName('a');

    for (let i = 0; i < folderLinks.length; i++) {
      const folderNode = folderLinks[i];
      const folderNodeUrl = folderNode.getAttribute('href')!;

      if (encodedFolderUrl !== folderNodeUrl) {
        continue;
      }

      const defaultValues = folderNode.getElementsByTagName('DefaultValue');

      for (let j = 0; j < defaultValues.length; j++) {
        const defaultValueNode = defaultValues[j];
        const defaultValueNodeField = defaultValueNode.getAttribute('FieldName')!;

        if (defaultValueNodeField !== fieldName) {
          continue;
        }

        // Default value node found, let's update the value
        defaultValueNode.textContent = fieldValue;
        return doc.toString();
      }

      // Default value node not found, let's create it
      folderNode.appendChild(newDefaultValueNode);
      return doc.toString();
    }

    // Folder node was not found, let's create it
    const newFolderNode = doc.createElement('a');
    newFolderNode.setAttribute('href', encodedFolderUrl);
    newFolderNode.appendChild(newDefaultValueNode);
    doc.documentElement!.appendChild(newFolderNode);

    return doc.toString();
  }

  private async uploadDefaultColumnValuesXml(logger: Logger, webUrl: string, listServerRelUrl: string, xml: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Uploading default column values to list...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listServerRelUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'text/plain'
      },
      responseType: 'json',
      data: xml
    };

    await request.put(requestOptions);
  }
}

export default new SpoListDefaultValueSetCommand();