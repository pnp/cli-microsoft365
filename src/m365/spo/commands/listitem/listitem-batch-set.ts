import fs from 'fs';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface FieldDetails {
  InternalName: string;
  TypeAsString: string;
}

interface UserDetail {
  email: string;
  id: number;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  idColumn?: string;
  systemUpdate?: boolean;
}

class SpoListItemBatchSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_BATCH_SET;
  }

  public get description(): string {
    return 'Updates list items in a batch';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        idColumn: typeof args.options.idColumn !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        systemUpdate: !!args.options.systemUpdate
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--idColumn [idColumn]'
      },
      {
        option: '-s, --systemUpdate'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (!fs.existsSync(args.options.filePath)) {
          return `File with path ${args.options.filePath} does not exist`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Starting to create batch items from csv at path ${args.options.filePath}`);
      }

      const csvContent = fs.readFileSync(args.options.filePath, 'utf8');
      const jsonContent: any[] = formatting.parseCsvToJson(csvContent);
      const amountOfRows = jsonContent.length;
      const idColumn = args.options.idColumn || 'ID';

      if (!jsonContent[0].hasOwnProperty(idColumn)) {
        throw `The specified value for idColumn does not exist in the array. Specified idColumn is '${args.options.idColumn || 'ID'}'. Please specify the correct value.`;
      }

      const listId = await this.getListId(args.options, logger);
      const fields = await this.getListFields(args.options, listId, jsonContent, idColumn, logger);
      const userFields = fields.filter(field => field.TypeAsString === 'UserMulti' || field.TypeAsString === 'User');
      const resolvedUsers = await this.getUsersFromCsv(args.options.webUrl, jsonContent, userFields);
      const formDigestValue = (await spo.getRequestDigest(args.options.webUrl)).FormDigestValue;
      const objectIdentity = (await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue)).objectIdentity;

      let counter = 0;

      while (jsonContent.length > 0) {
        const entriesToProcess = jsonContent.splice(0, 50);
        const objectPaths = [], actions = [];
        let index = 1;

        for (const row of entriesToProcess) {
          counter += 1;
          objectPaths.push(`<Identity Id="${index}" Name="${objectIdentity}:list:${listId}:item:${row[idColumn]},1" />`);
          const [actionString, updatedIndex] = this.mapActions(index, row, fields, resolvedUsers, args.options.systemUpdate);
          index = updatedIndex;
          actions.push(actionString);
        }

        if (this.verbose) {
          await logger.logToStderr(`Writing away batch of items, currently at: ${counter}/${amountOfRows}.`);
        }

        await this.sendBatchRequest(args.options.webUrl, this.getRequestBody(objectPaths, actions));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestBody(objectPaths: string[], actions: string[]): string {
    return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${actions.join('')}</Actions><ObjectPaths>${objectPaths.join('')}</ObjectPaths></Request>`;
  }

  private async sendBatchRequest(webUrl: string, requestBody: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'Content-Type': 'text/xml'
      },
      data: requestBody
    };
    const res: any = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage + ' - ' + response.ErrorInfo.ErrorValue;
    }
  }

  private mapActions(index: number, row: any, fields: FieldDetails[], users: UserDetail[], systemUpdate?: boolean): [string, number] {
    const objectPathId = index;
    let actionString = '';

    fields.forEach((field: FieldDetails) => {

      if (row[field.InternalName] === undefined || row[field.InternalName] === '') {
        actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String"></Parameter></Parameters></Method>`;
      }
      else {
        switch (field.TypeAsString) {
          case 'User':
            const userDetail = users.find(us => us.email === row[field.InternalName])!;
            actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String">${userDetail.id}</Parameter></Parameters></Method>`;
            break;
          case 'UserMulti':
            const userMultiString: string[] = row[field.InternalName].toString().split(';').map((element: string) => {
              const userDetail = users.find(us => us.email === element)!;
              return `<Object TypeId="{c956ab54-16bd-4c18-89d2-996f57282a6f}"><Property Name="Email" Type="Null" /><Property Name="LookupId" Type="Int32">${userDetail.id}</Property><Property Name="LookupValue" Type="Null" /></Object>`;
            });
            actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="Array">${userMultiString.join('')}</Parameter></Parameters></Method>`;
            break;
          case 'Lookup':
            actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter TypeId="{f1d34cc0-9b50-4a78-be78-d5facfcccfb7}"><Property Name="LookupId" Type="Int32">${row[field.InternalName]}</Property><Property Name="LookupValue" Type="Null"/></Parameter></Parameters></Method>`;
            break;
          case 'LookupMulti':
            const lookupMultiString: string[] = row[field.InternalName].toString().split(';').map((element: string) => {
              return `<Object TypeId="{f1d34cc0-9b50-4a78-be78-d5facfcccfb7}"><Property Name="LookupId" Type="Int32">${element}</Property><Property Name="LookupValue" Type="Null" /></Object>`;
            });
            actionString += `<Method Name="SetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="Array">${lookupMultiString.join('')}</Parameter></Parameters></Method>`;
            break;
          default:
            actionString += `<Method Name="ParseAndSetFieldValue" Id="${index += 1}" ObjectPathId="${objectPathId}"><Parameters><Parameter Type="String">${field.InternalName}</Parameter><Parameter Type="String">${(<any>row)[field.InternalName].toString()}</Parameter></Parameters></Method>`;
            break;
        }
      }
    });

    actionString += `<Method Name="${systemUpdate ? 'System' : ''}Update" Id="${index += 1}" ObjectPathId="${objectPathId}"/>`;
    return [actionString, index];
  }

  private async getListFields(options: Options, listId: string, jsonContent: any, idColumn: string, logger: Logger): Promise<FieldDetails[]> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving fields for list with id ${listId}`);
    }

    const filterFields: string[] = [];
    const objectKeys = Object.keys(jsonContent[0]);
    const index = objectKeys.indexOf(idColumn, 0);

    if (index > -1) {
      objectKeys.splice(index, 1);
    }

    objectKeys.map(objectKey => {
      filterFields.push(`InternalName eq '${objectKey}'`);
    });

    const fields = await odata.getAllItems<FieldDetails>(`${options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`);

    if (fields.length !== objectKeys.length) {
      const fieldsThatDontExist: string[] = [];

      objectKeys.forEach(key => {
        const field = fields.find(field => field.InternalName === key);

        if (!field) {
          fieldsThatDontExist.push(key);
        }
      });

      throw `Following fields specified in the csv do not exist on the list: ${fieldsThatDontExist.join(', ')}`;
    }

    return fields;
  }

  private async getListId(options: Options, logger: Logger): Promise<string> {
    if (options.listId) {
      return options.listId;
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving list id');
    }

    let listUrl = `${options.webUrl}/_api/web`;

    if (options.listTitle) {
      listUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl!);
      listUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${listUrl}?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listResult = await request.get<{ Id: string }>(requestOptions);
    return listResult.Id;
  }

  private async getUsersFromCsv(webUrl: string, jsonContent: any[], userFields: FieldDetails[]): Promise<UserDetail[]> {
    if (userFields.length === 0) {
      return [];
    }

    const userFieldValues: UserDetail[] = [];
    const emailsToResolve = this.getEmailsToEnsure(jsonContent, userFields);

    for await (const email of emailsToResolve) {
      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/web/ensureUser('${email}')?$select=Id`,
        headers: {
          accept: 'application/json',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      const response = await request.post<{ Id: number }>(requestOptions);
      userFieldValues.push({ email: email, id: response.Id });
    }

    return userFieldValues;
  }

  private getEmailsToEnsure(jsonContent: any[], userFields: FieldDetails[]): string[] {
    const emailsToResolve: string[] = [];

    userFields.forEach((userField: FieldDetails) => {

      jsonContent.forEach(row => {
        const fieldValue = row[userField.InternalName];

        if (fieldValue !== undefined && fieldValue !== '') {
          const emailsSplitted = fieldValue.split(';');
          emailsSplitted.forEach((email: string) => {
            if (!emailsToResolve.some(existingMail => existingMail === email)) {
              emailsToResolve.push(email);
            }
          });
        }
      });
    });

    return emailsToResolve;
  }
}

export default new SpoListItemBatchSetCommand();