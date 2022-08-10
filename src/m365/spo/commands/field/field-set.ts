import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  name?: string;
  listId?: string;
  listTitle?: string;
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
        name: typeof args.options.name !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
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
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
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

	    if (args.options.listId && args.options.listTitle) {
	      return `Specify listId or listTitle but not both`;
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
  	this.optionSets.push(['id', 'title', 'name']);
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (args.options.name) {
      this.warn(logger, `Option 'name' is deprecated. Please use 'title' instead.`);
    }

    let requestDigest: string = '';

    spo
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<string> => {
        requestDigest = res.FormDigestValue;

        if (!args.options.listId && !args.options.listTitle) {
          return Promise.resolve(undefined as any);
        }

        const listQuery: string = args.options.listId ?
          `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.listId)}</Parameter></Parameters></Method>` :
          `<Method Id="663" ParentId="7" Name="GetByTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.listTitle)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${listQuery}<Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res?: string): Promise<string> => {
        // by default retrieve the column from the site
        let fieldsParentIdentity: string = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';

        if (res) {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            return Promise.reject(response.ErrorInfo.ErrorMessage);
          }

          const result: { _ObjectIdentity_: string; } = json[json.length - 1];
          fieldsParentIdentity = `<Identity Id="5" Name="${result._ObjectIdentity_}" />`;
        }

        // retrieve column CSOM object id
        const fieldQuery: string = args.options.id ?
          `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method>` :
          `<Method Id="663" ParentId="7" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name || args.options.title)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${fieldQuery}<Property Id="7" ParentId="5" Name="Fields" />${fieldsParentIdentity}</ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<string> => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          return Promise.reject(response.ErrorInfo.ErrorMessage);
        }

        const result: { _ObjectIdentity_: string; } = json[json.length - 1];
        const fieldId: string = result._ObjectIdentity_;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${this.getPayload(args.options)}<Method Name="UpdateAndPushChanges" Id="9000" ObjectPathId="663"><Parameters><Parameter Type="Boolean">${args.options.updateExistingLists ? 'true' : 'false'}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="663" Name="${fieldId}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getPayload(options: any): string {
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'id',
      'name',
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