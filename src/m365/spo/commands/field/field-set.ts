import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
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
  id?: string;
  title?: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  updateExistingLists?: boolean;
  webUrl: string;
}

class SpoFieldSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_SET;
  }

  public get description(): string {
    return 'Updates existing list or site column';
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
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        updateExistingLists: !!args.options.updateExistingLists
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--updateExistingLists'
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

        const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either list id or title or list url, but not multiple`;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (args.options.id &&
          !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} in option id is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const requestDigest = reqDigest.FormDigestValue;

      let fieldsParentIdentity = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';

      if (args.options.listId || args.options.listTitle || args.options.listUrl) {
        let requestData = '';
        if (args.options.listId) {
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.listId)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
        }
        else if (args.options.listTitle) {
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetByTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.listTitle)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
        }
        else if (args.options.listUrl) {
          const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="3" ParentId="1" Name="Web" /><Method Id="5" ParentId="3" Name="GetList"><Parameters><Parameter Type="String">${listServerRelativeUrl}</Parameter></Parameters></Method></ObjectPaths></Request>`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: requestData
        };

        const list = await request.post<string>(requestOptions);
        const json: ClientSvcResponse = JSON.parse(list);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }

        const result: { _ObjectIdentity_: string; } = json[json.length - 1];
        fieldsParentIdentity = `<Identity Id="5" Name="${result._ObjectIdentity_}" />`;
      }

      // retrieve column CSOM object id
      const fieldQuery: string = args.options.id ?
        `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method>` :
        `<Method Id="663" ParentId="7" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name || args.options.title)}</Parameter></Parameters></Method>`;

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': requestDigest
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${fieldQuery}<Property Id="7" ParentId="5" Name="Fields" />${fieldsParentIdentity}</ObjectPaths></Request>`
      };

      const field = await request.post<string>(requestOptions);
      let json: ClientSvcResponse = JSON.parse(field);
      let response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const result: { _ObjectIdentity_: string; } = json[json.length - 1];
      const fieldId: string = result._ObjectIdentity_;

      requestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': requestDigest
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${this.getPayload(args.options)}<Method Name="UpdateAndPushChanges" Id="9000" ObjectPathId="663"><Parameters><Parameter Type="Boolean">${args.options.updateExistingLists ? 'true' : 'false'}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="663" Name="${fieldId}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      json = JSON.parse(res);
      response = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private getPayload(options: any): string {
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'listUrl',
      'id',
      'title',
      'updateExistingLists',
      'debug',
      'verbose',
      'output'
    ];

    let i: number = 667;
    const payload: string = Object.keys(options).map(key => {
      return excludeOptions.indexOf(key) === -1 ? `<SetProperty Id="${i++}" ObjectPathId="663" Name="${key}"><Parameter Type="String">${formatting.escapeXml(options[key])}</Parameter></SetProperty>` : '';
    }).join('');

    return payload;
  }
}

module.exports = new SpoFieldSetCommand();