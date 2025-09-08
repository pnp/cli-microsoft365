import os from 'os';
import { urlUtil } from "./urlUtil.js";
import { Logger } from "../cli/Logger.js";
import request, { CliRequestOptions } from "../request.js";
import { formatting } from './formatting.js';
import { odata, ODataResponse } from './odata.js';
import { ListItemInstance } from '../m365/spo/commands/listitem/ListItemInstance.js';
import { ListItemFieldValueResult } from '../m365/spo/commands/listitem/ListItemFieldValueResult.js';
import { ListItemInstanceCollection } from '../m365/spo/commands/listitem/ListItemInstanceCollection.js';
import { spo } from './spo.js';
import { basic } from './basic.js';

interface ContentType {
  Id: {
    StringValue: string;
  };
  Name: string;
}
interface ListSelectionOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

export interface ListItemListOptions extends ListSelectionOptions {
  fields?: string[];
  filter?: string;
  pageNumber?: number;
  pageSize?: number;
  camlQuery?: string;
  webUrl: string;
}

export interface ListItemAddOptions extends ListSelectionOptions {
  contentType?: string;
  folder?: string;
  fieldValues: { [key: string]: any };
}

function getListApiUrl(options: ListSelectionOptions): string {
  let listApiUrl = `${options.webUrl}/_api/web`;

  if (options.listId) {
    listApiUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
  }
  else if (options.listTitle) {
    listApiUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
  }
  else if (options.listUrl) {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
    listApiUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
  }

  return listApiUrl;
};

function getExpandFieldsArray(fieldsArray: string[]): string[] {
  const fieldsWithSlash: string[] = fieldsArray.filter(item => item.includes('/'));
  const fieldsToExpand: string[] = fieldsWithSlash.map(e => e.split('/')[0]);
  const expandFieldsArray: string[] = fieldsToExpand.filter((item, pos) => fieldsToExpand.indexOf(item) === pos);
  return expandFieldsArray;
}

async function getLastItemIdForPage(options: ListItemListOptions, listApiUrl: string, logger: Logger, verbose: boolean): Promise<number | undefined> {
  if (!(options.pageNumber) || Number(options.pageNumber) === 0 || !(options.pageSize)) {
    return undefined;
  }

  if (verbose) {
    await logger.logToStderr(`Getting skipToken Id for page ${options.pageNumber}`);
  }

  const rowLimit: string = `$top=${Number(options.pageSize) * Number(options.pageNumber)}`;
  const filter: string = options.filter ? `$filter=${encodeURIComponent(options.filter)}` : ``;

  const requestOptions: CliRequestOptions = {
    url: `${listApiUrl}/items?$select=Id&${rowLimit}&${filter}`,
    headers: {
      'accept': 'application/json;odata=nometadata'
    },
    responseType: 'json'
  };

  const response = await request.get<{ value: [{ Id: number }] }>(requestOptions);
  return response.value[response.value.length - 1]?.Id;
}

async function getListItemsByCamlQuery(options: ListItemListOptions, listApiUrl: string, logger: Logger, verbose: boolean): Promise<ListItemInstance[]> {
  const formDigestValue = (await spo.getRequestDigest(options.webUrl)).FormDigestValue;

  if (verbose) {
    await logger.logToStderr(`Getting list items using CAML query`);
  }

  const items: ListItemInstance[] = [];
  let skipTokenId: number | undefined = undefined;

  do {
    const requestBody: any = {
      "query": {
        "ViewXml": options.camlQuery,
        "AllowIncrementalResults": true
      }
    };

    if (skipTokenId !== undefined) {
      requestBody.query.ListItemCollectionPosition = {
        "PagingInfo": `Paged=TRUE&p_ID=${skipTokenId}`
      };
    }

    const requestOptions: CliRequestOptions = {
      url: `${listApiUrl}/GetItems`,
      headers: {
        'accept': 'application/json;odata=nometadata',
        'X-RequestDigest': formDigestValue
      },
      responseType: 'json',
      data: requestBody
    };

    const listItemInstances = await request.post<ListItemInstanceCollection>(requestOptions);
    skipTokenId = listItemInstances.value.length > 0 ? listItemInstances.value[listItemInstances.value.length - 1].Id : undefined;
    items.push(...listItemInstances.value);
  }
  while (skipTokenId !== undefined);

  return items;
}

async function getListItems(options: ListItemListOptions, listApiUrl: string, logger: Logger, verbose: boolean): Promise<ListItemInstance[]> {
  if (verbose) {
    await logger.logToStderr(`Getting list items`);
  }

  const fieldsArray: string[] = options.fields ? options.fields : [];
  const expandFieldsArray: string[] = getExpandFieldsArray(fieldsArray);
  const queryParams = [options.pageSize ? `$top=${options.pageSize}` : '$top=5000'];
  const skipTokenId = await getLastItemIdForPage(options, listApiUrl, logger, verbose);

  if (options.filter) {
    queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
  }

  if (expandFieldsArray.length > 0) {
    queryParams.push(`$expand=${expandFieldsArray.join(",")}`);
  }

  if (fieldsArray.length > 0) {
    queryParams.push(`$select=${formatting.encodeQueryParameter(fieldsArray.join(','))}`);
  }

  if (skipTokenId !== undefined) {
    queryParams.push(`$skiptoken=Paged=TRUE%26p_ID=${skipTokenId}`);
  }

  // If skiptoken is not found, then we are past the last page
  if (options.pageNumber && Number(options.pageNumber) > 0 && skipTokenId === undefined) {
    return [];
  }

  if (!options.pageSize) {
    return await odata.getAllItems<ListItemInstance>(`${listApiUrl}/items?${queryParams.join('&')}`);
  }

  const requestOptions: CliRequestOptions = {
    url: `${listApiUrl}/items?${queryParams.join('&')}`,
    headers: {
      'accept': 'application/json;odata=nometadata'
    },
    responseType: 'json'
  };

  const listItemCollection = await request.get<ListItemInstanceCollection>(requestOptions);
  return listItemCollection.value;
}

async function getContentTypeName(options: ListItemAddOptions, listApiUrl: string, logger: Logger, verbose: boolean, debug: boolean): Promise<string | undefined> {
  if (!options.contentType) {
    return undefined;
  }

  let contentTypeName: string = '';

  if (verbose) {
    await logger.logToStderr(`Getting content types for list...`);
  }

  const requestOptions: CliRequestOptions = {
    url: `${listApiUrl}/contenttypes?$select=Name,Id`,
    headers: {
      'accept': 'application/json;odata=nometadata'
    },
    responseType: 'json'
  };

  const contentTypes = await request.get<ODataResponse<ContentType>>(requestOptions);
  const foundContentType = await basic.asyncFilter<ContentType>(contentTypes.value, async (ct: ContentType) => {
    const contentTypeMatch: boolean = ct.Id.StringValue === options.contentType || ct.Name === options.contentType;

    if (debug) {
      await logger.logToStderr(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
    }

    return contentTypeMatch;
  });

  if (debug) {
    await logger.logToStderr('Content type filter output...');
    await logger.logToStderr(foundContentType);
  }

  if (foundContentType.length > 0) {
    contentTypeName = foundContentType[0].Name;
  }

  // After checking for content types, throw an error if the name is blank
  if (!contentTypeName || contentTypeName === '') {
    throw new Error(`Specified content type '${options.contentType}' doesn't exist on the target list`);
  }

  if (debug) {
    await logger.logToStderr(`Using content type name: ${contentTypeName}`);
  }

  return contentTypeName;
}

async function ensureTargetFolder(options: ListItemAddOptions, listApiUrl: string, logger: Logger, verbose: boolean, debug: boolean): Promise<string | undefined> {
  if (!options.folder) {
    return undefined;
  }

  if (verbose) {
    await logger.logToStderr('Setting up folder lookup response ...');
  }

  const requestOptions: CliRequestOptions = {
    url: `${listApiUrl}/rootFolder`,
    headers: {
      'accept': 'application/json;odata=nometadata'
    },
    responseType: 'json'
  };

  const rootFolderResponse = await request.get<any>(requestOptions);
  const targetFolderServerRelativeUrl = urlUtil.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], options.folder as string);
  await spo.ensureFolder(options.webUrl, targetFolderServerRelativeUrl, logger, debug === true);

  return targetFolderServerRelativeUrl;
}

function mapListItemCreationRequestBody(options: ListItemAddOptions): any {
  const requestBody: any = [];

  Object.keys(options.fieldValues).forEach(key => {
    requestBody.push({ FieldName: key, FieldValue: `${options.fieldValues[key]}` });
  });

  return requestBody;
}

export const spoListItem = {
  /**
  * Get the listitems of a SharePoint list.
  * Returns an array of ListItemInstance or an array with the field properties if supplied
  * @param options The options to get the list items
  * @param logger The logger object
  * @param verbose If the function is executed in verbose mode
  */
  async getListItems(options: ListItemListOptions, logger: Logger, verbose: boolean): Promise<ListItemInstance[]> {
    const listApiUrl = getListApiUrl(options);

    const listItems = options.camlQuery ?
      await getListItemsByCamlQuery(options, listApiUrl, logger, verbose) :
      await getListItems(options, listApiUrl, logger, verbose);

    listItems.forEach(v => delete v['ID']);

    return listItems;
  },

  /**
  * Adds a list item to a list
  * Returns a ListItemInstance object
  * @param options The options relting to then new listitem
  * @param logger the Logger object
  * @param verbose If the function is executed in verbose mode
  * @param debug If the function is executed in debug mode
  */
  async addListItem(options: ListItemAddOptions, logger: Logger, verbose: boolean, debug: boolean): Promise<ListItemInstance> {
    const listApiUrl = getListApiUrl(options);

    const contentTypeName: string | undefined = await getContentTypeName(options, listApiUrl, logger, verbose, debug);
    const targetFolderServerRelativeUrl: string | undefined = await ensureTargetFolder(options, listApiUrl, logger, verbose, debug);

    const requestBody: any = {
      formValues: mapListItemCreationRequestBody(options)
    };

    if (options.folder) {
      requestBody.listItemCreateInfo = {
        FolderPath: {
          DecodedUrl: targetFolderServerRelativeUrl
        }
      };
    }

    if (options.contentType && contentTypeName !== undefined) {
      if (debug) {
        await logger.logToStderr(`Specifying content type name [${contentTypeName}] in request body`);
      }

      requestBody.formValues.push({
        FieldName: 'ContentType',
        FieldValue: contentTypeName
      });
    }

    if (verbose) {
      await logger.logToStderr(`Adding a list item in list '${options.listId || options.listTitle || options.listUrl}'...`);
    }

    const postRequestOptions: CliRequestOptions = {
      url: `${listApiUrl}/AddValidateUpdateItemUsingPath()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    const response = await request.post<any>(postRequestOptions);

    // Response is from /AddValidateUpdateItemUsingPath POST call, perform get on added item to get all field values
    const fieldValues: ListItemFieldValueResult[] = response.value;
    if (fieldValues.some(f => f.HasException)) {
      throw new Error(`Creating the item failed with the following errors: ${os.EOL}${fieldValues.filter(f => f.HasException).map(f => { return `- ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`);
    }

    const idField = fieldValues.filter((thisField) => {
      return (thisField.FieldName === "Id");
    });

    if (debug) {
      await logger.logToStderr(`Field values returned:`);
      await logger.logToStderr(fieldValues);
      await logger.logToStderr(`Id returned by AddValidateUpdateItemUsingPath: ${idField[0].FieldValue}`);
    }

    const getRequestOptions: CliRequestOptions = {
      url: `${listApiUrl}/items(${idField[0].FieldValue})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const item = await request.get<ListItemInstance>(getRequestOptions);
    delete item.ID;

    return item;
  }
};