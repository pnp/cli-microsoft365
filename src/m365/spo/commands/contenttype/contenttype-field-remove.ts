import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  fieldLinkId: string;
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  updateChildContentTypes?: boolean;
  confirm?: boolean;
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes a column from a site- or list content type';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        updateChildContentTypes: !!args.options.updateChildContentTypes,
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --contentTypeId <contentTypeId>'
      },
      {
        option: '-f, --fieldLinkId <fieldLinkId>'
      },
      {
        option: '-c, --updateChildContentTypes'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.fieldLinkId)) {
          return `${args.options.fieldLinkId} is not a valid GUID`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('i', 'contentTypeId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeFieldLink: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.debug) {
          logger.logToStderr(`Get SiteId required by ProcessQuery endpoint.`);
        }

        // GET SiteId
        let requestOptions: AxiosRequestConfig = {
          url: `${args.options.webUrl}/_api/site?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const site = await request.get<{ Id: string }>(requestOptions);
        const siteId = site.Id;

        if (this.debug) {
          logger.logToStderr(`SiteId: ${siteId}`);
          logger.logToStderr(`Get WebId required by ProcessQuery endpoint.`);
        }

        // GET WebId
        requestOptions = {
          url: `${args.options.webUrl}/_api/web?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const web = await request.get<{ Id: string }>(requestOptions);
        const webId = web.Id;

        if (this.debug) {
          logger.logToStderr(`WebId: ${webId}`);
        }

        let listId: string | undefined = undefined;

        if (args.options.listId) {
          listId = args.options.listId;
        }
        else if (args.options.listTitle) {
          listId = await this.getListIdFromListTitle(args.options.webUrl, args.options.listTitle);
        }
        else if (args.options.listUrl) {
          listId = await this.getListIdFromListUrl(args.options.webUrl, args.options.listUrl);
        }

        if (this.debug) {
          logger.logToStderr(`ListId: ${listId}`);
        }

        const reqDigest = await spo.getRequestDigest(args.options.webUrl);
        const requestDigest: string = reqDigest.FormDigestValue;

        const updateChildContentTypes: boolean = args.options.listTitle || args.options.listId || args.options.listUrl ? false : args.options.updateChildContentTypes === true;

        if (this.debug) {
          const additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : args.options.listId ? `; ListId='${args.options.listId}'` : args.options.listUrl ? `; ListUrl='${args.options.listUrl}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
          logger.logToStderr(`Remove FieldLink from ContentType. FieldLinkId='${args.options.fieldLinkId}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
          logger.logToStderr(`Execute ProcessQuery.`);
        }

        let requestBody: string = '';
        if (listId) {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
        }
        else {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
        }

        requestOptions = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: requestBody
        };

        const res = await request.post<string>(requestOptions);
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeFieldLink();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the column ${args.options.fieldLinkId} from content type ${args.options.contentTypeId}?`
      });

      if (result.continue) {
        await removeFieldLink();
      }
    }
  }

  private async getListIdFromListTitle(webUrl: string, listTitle: string): Promise<string> {
    const requestOptions: AxiosRequestConfig = {
      url: `${webUrl}/_api/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const list = await request.get<{ Id: string }>(requestOptions);
    return list.Id;
  }

  private async getListIdFromListUrl(webUrl: string, listUrl: string): Promise<string> {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
    const requestOptions: AxiosRequestConfig = {
      url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const list = await request.get<{ Id: string }>(requestOptions);
    return list.Id;
  }
}

module.exports = new SpoContentTypeFieldRemoveCommand();