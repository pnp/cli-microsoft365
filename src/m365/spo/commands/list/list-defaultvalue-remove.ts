import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { DOMParser } from '@xmldom/xmldom';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';

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
  folderUrl: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface RemoveDefaultValueResult {
  isFieldFound: boolean;
  xml?: string;
}

class SpoListDefaultValueRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_DEFAULTVALUE_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific default column value for a specific document library';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: listId, listTitle, listUrl.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove default column value '${args.options.fieldName}' from ${args.options.folderUrl ? `'${args.options.folderUrl}'` : 'the root of the list'}?` });

      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Removing default column value '${args.options.fieldName}' from ${args.options.folderUrl ? `'${args.options.folderUrl}'` : 'the root of the list'}.`);
        await logger.logToStderr(`Getting server-relative URL of the list...`);
      }

      const listServerRelUrl = await this.getServerRelativeListUrl(args.options);
      let folderUrl = listServerRelUrl;

      if (args.options.folderUrl) {
        folderUrl = urlUtil.getServerRelativePath(args.options.webUrl, urlUtil.removeTrailingSlashes(args.options.folderUrl));
      }

      if (this.verbose) {
        await logger.logToStderr(`Getting default column values...`);
      }

      const defaultValuesXml = await this.getDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl);
      const removeDefaultValueResult = this.removeFieldFromXml(defaultValuesXml, args.options.fieldName, folderUrl);

      if (!removeDefaultValueResult.isFieldFound) {
        throw `Default column value '${args.options.fieldName}' was not found.`;
      }

      if (this.verbose) {
        await logger.logToStderr(`Uploading default column values to list...`);
      }

      await this.uploadDefaultColumnValuesXml(args.options.webUrl, listServerRelUrl, removeDefaultValueResult.xml!);
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

  private removeFieldFromXml(xml: string | null, fieldName: string, folderUrl: string): RemoveDefaultValueResult {
    if (xml === null) {
      return { isFieldFound: false };
    }

    // Encode all spaces in the folder URL
    const encodedFolderUrl = folderUrl.replace(/ /g, '%20');

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    const folderLinks = doc.getElementsByTagName('a');

    for (let i = 0; i < folderLinks.length; i++) {
      const folderNode = folderLinks[i];
      const folderNodeUrl = folderNode.getAttribute('href')!;

      if (encodedFolderUrl.toLowerCase() !== folderNodeUrl.toLowerCase()) {
        continue;
      }

      const defaultValues = folderNode.getElementsByTagName('DefaultValue');

      for (let j = 0; j < defaultValues.length; j++) {
        const defaultValueNode = defaultValues[j];
        const defaultValueNodeField = defaultValueNode.getAttribute('FieldName')!;

        if (defaultValueNodeField !== fieldName) {
          continue;
        }

        if (folderNode.childNodes.length === 1) {
          // No other default values found in the folder, let's remove the folder node
          folderNode.parentNode!.removeChild(folderNode);
        }
        else {
          // Default value node found, let's remove it
          folderNode.removeChild(defaultValueNode);
        }
        return { isFieldFound: true, xml: doc.toString() };
      }
    }

    return { isFieldFound: false };
  }

  private async uploadDefaultColumnValuesXml(webUrl: string, listServerRelUrl: string, xml: string): Promise<void> {
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

export default new SpoListDefaultValueRemoveCommand();