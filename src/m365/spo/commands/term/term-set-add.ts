import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import { v4 } from 'uuid';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandError,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { TermSet } from './TermSet';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  customProperties?: string;
  description?: string;
  id?: string;
  name: string;
  termGroupId?: string;
  termGroupName?: string;
}

class SpoTermSetAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_SET_ADD;
  }

  public get description(): string {
    return 'Adds taxonomy term set';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.customProperties = typeof args.options.customProperties !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.termGroupId = typeof args.options.termGroupId !== 'undefined';
    telemetryProps.termGroupName = typeof args.options.termGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let formDigest: string = '';
    let termSet: TermSet;
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        formDigest = res.FormDigestValue;

        if (this.verbose) {
          cmd.log(`Adding taxonomy term set...`);
        }

        const termGroupQuery: string = args.options.termGroupName ?
          `<Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>` :
          `<Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>`;
        const termSetId: string = args.options.id || v4();

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" />${termGroupQuery}<Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter><Parameter Type="Guid">{${termSetId}}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<string> => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          return Promise.reject(response.ErrorInfo.ErrorMessage);
        }

        termSet = json[json.length - 1];

        if (!args.options.description &&
          !args.options.customProperties) {
          return Promise.resolve(undefined as any);
        }

        let termStoreObjectIdentity: string = '';
        // get term store object identity
        for (let i: number = 0; i < json.length; i++) {
          if (json[i] !== 39) {
            continue;
          }

          termStoreObjectIdentity = json[i + 1]._ObjectIdentity_;
          break;
        }

        if (this.verbose) {
          cmd.log(`Setting term set properties...`);
        }

        const properties: string[] = [];
        let i: number = 127;
        if (args.options.description) {
          properties.push(`<SetProperty Id="${i++}" ObjectPathId="117" Name="Description"><Parameter Type="String">${Utils.escapeXml(args.options.description)}</Parameter></SetProperty>`);
          termSet.Description = args.options.description;
        }
        if (args.options.customProperties) {
          const customProperties: any = JSON.parse(args.options.customProperties);
          Object.keys(customProperties).forEach(k => {
            properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${Utils.escapeXml(k)}</Parameter><Parameter Type="String">${Utils.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
          });
          termSet.CustomProperties = customProperties;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigest
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${termSet._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res?: string): void => {
        if (res) {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }
        }

        delete termSet._ObjectIdentity_;
        delete termSet._ObjectType_;
        termSet.CreatedDate = new Date(Number(termSet.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        termSet.Id = termSet.Id.replace('/Guid(', '').replace(')/', '');
        termSet.LastModifiedDate = new Date(Number(termSet.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        cmd.log(termSet);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the term set to add'
      },
      {
        option: '--termGroupId [termGroupId]',
        description: 'ID of the term group in which to create the term set. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termGroupName [termGroupName]',
        description: 'Name of the term group in which to create the term set. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the term set to add'
      },
      {
        option: '-d, --description [description]',
        description: 'Description of the term set to add'
      },
      {
        option: '--customProperties [customProperties]',
        description: 'JSON string with key-value pairs representing custom properties to set on the term set'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (!args.options.termGroupId && !args.options.termGroupName) {
        return 'Specify termGroupId or termGroupName';
      }

      if (args.options.termGroupId && args.options.termGroupName) {
        return 'Specify termGroupId or termGroupName but not both';
      }

      if (args.options.termGroupId) {
        if (!Utils.isValidGuid(args.options.termGroupId)) {
          return `${args.options.termGroupId} is not a valid GUID`;
        }
      }

      if (args.options.customProperties) {
        try {
          JSON.parse(args.options.customProperties);
        }
        catch (e) {
          return `Error when parsing customProperties JSON: ${e}`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoTermSetAddCommand();