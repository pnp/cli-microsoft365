import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
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
  id?: string;
  name?: string;
  termGroupId?: string;
  termGroupName?: string;
}

class SpoTermSetGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_SET_GET;
  }

  public get description(): string {
    return 'Gets information about the specified taxonomy term set';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.termGroupId = typeof args.options.termGroupId !== 'undefined';
    telemetryProps.termGroupName = typeof args.options.termGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Retrieving taxonomy term set...`);
        }

        const termGroupQuery: string = args.options.termGroupId ? `<Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
        const termSetQuery: string = args.options.id ? `<Method Id="67" ParentId="65" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method>` : `<Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" />${termGroupQuery}<Property Id="65" ParentId="62" Name="TermSets" />${termSetQuery}</ObjectPaths></Request>`
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

        const termSet: TermSet = json[json.length - 1];
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
        option: '-i, --id [id]',
        description: 'ID of the term set to retrieve. Specify name or id but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the term set to retrieve. Specify name or id but not both'
      },
      {
        option: '--termGroupId [termGroupId]',
        description: 'ID of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termGroupName [termGroupName]',
        description: 'Name of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id && !args.options.name) {
        return 'Specify either id or name';
      }

      if (args.options.id && args.options.name) {
        return 'Specify either id or name but not both';
      }

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

      return true;
    };
  }
}

module.exports = new SpoTermSetGetCommand();