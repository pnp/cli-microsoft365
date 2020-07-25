import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.updateExistingLists = !!args.options.updateExistingLists;
    return telemetryProps;
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let requestDigest: string = '';

    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<string> => {
        requestDigest = res.FormDigestValue;

        if (!args.options.listId && !args.options.listTitle) {
          return Promise.resolve(undefined as any);
        }

        const listQuery: string = args.options.listId ?
          `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${Utils.escapeXml(args.options.listId)}</Parameter></Parameters></Method>` :
          `<Method Id="663" ParentId="7" Name="GetByTitle"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.listTitle)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${listQuery}<Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
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
          `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${Utils.escapeXml(args.options.id)}</Parameter></Parameters></Method>` :
          `<Method Id="663" ParentId="7" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${fieldQuery}<Property Id="7" ParentId="5" Name="Fields" />${fieldsParentIdentity}</ObjectPaths></Request>`
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${this.getPayload(args.options)}<Method Name="UpdateAndPushChanges" Id="9000" ObjectPathId="663"><Parameters><Parameter Type="Boolean">${args.options.updateExistingLists ? 'true' : 'false'}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="663" Name="${fieldId}" /></ObjectPaths></Request>`
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

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getPayload(options: any): string {
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'id',
      'name',
      'updateExistingLists',
      'debug',
      'verbose',
      'output'
    ];

    let i: number = 667;
    const payload: string = Object.keys(options).map(key => {
      return excludeOptions.indexOf(key) === -1 ? `<SetProperty Id="${i++}" ObjectPathId="663" Name="${key}"><Parameter Type="String">${Utils.escapeXml(options[key])}</Parameter></SetProperty>` : '';
    }).join('');

    return payload;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the field is located'
      },
      {
        option: '--listId [listId]',
        description: 'ID of the list where the field is located (if list column). Specify listTitle or listId but not both'
      },
      {
        option: '--listTitle [listTitle]',
        description: 'Title of the list where the field is located (if list column). Specify listTitle or listId but not both'
      },
      {
        option: '-i|--id [id]',
        description: 'ID of the field to update. Specify name or id but not both'
      },
      {
        option: '-n|--name [name]',
        description: 'Title or internal name of the field to update. Specify name or id but not both'
      },
      {
        option: '--updateExistingLists',
        description: 'Set, to push the update to existing lists. Otherwise, the changes will apply to new lists only'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (!args.options.id && !args.options.name) {
        return `Specify id or name`;
      }

      if (args.options.id && args.options.name) {
        return `Specify viewId or viewTitle but not both`;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} in option id is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoFieldSetCommand();