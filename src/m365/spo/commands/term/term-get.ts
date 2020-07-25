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
import { Term } from './Term';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  termGroupId?: string;
  termGroupName?: string;
  termSetId?: string;
  termSetName?: string;
}

class SpoTermGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_GET;
  }

  public get description(): string {
    return 'Gets information about the specified taxonomy term';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.termGroupId = typeof args.options.termGroupId !== 'undefined';
    telemetryProps.termGroupName = typeof args.options.termGroupName !== 'undefined';
    telemetryProps.termSetId = typeof args.options.termSetId !== 'undefined';
    telemetryProps.termSetName = typeof args.options.termSetName !== 'undefined';
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
          cmd.log(`Retrieving taxonomy term...`);
        }

        let body: string = '';

        if (args.options.id) {
          body = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method></ObjectPaths></Request>`
        }
        else {
          const termGroupQuery: string = args.options.termGroupId ? `<Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
          const termSetQuery: string = args.options.termSetId ? `<Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
          body = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" />${termGroupQuery}<Property Id="101" ParentId="98" Name="TermSets" />${termSetQuery}<Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: body
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

        const term: Term | null = json[json.length - 1];
        if (!term) {
          cb();
          return;
        }

        delete term._ObjectIdentity_;
        delete term._ObjectType_;
        term.CreatedDate = new Date(Number(term.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
        term.LastModifiedDate = new Date(Number(term.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        cmd.log(term);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the term to retrieve. Specify name or id but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the term to retrieve. Specify name or id but not both'
      },
      {
        option: '--termGroupId [termGroupId]',
        description: 'ID of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termGroupName [termGroupName]',
        description: 'Name of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termSetId [termSetId]',
        description: 'ID of the term set to which the term belongs. Specify termSetId or termSetName but not both'
      },
      {
        option: '--termSetName [termSetName]',
        description: 'Name of the term set to which the term belongs. Specify termSetId or termSetName but not both'
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

      if (args.options.name) {
        if (!args.options.termGroupId && !args.options.termGroupName) {
          return 'Specify termGroupId or termGroupName';
        }

        if (!args.options.termSetId && !args.options.termSetName) {
          return 'Specify termSetId or termSetName';
        }
      }

      if (args.options.termGroupId && args.options.termGroupName) {
        return 'Specify termGroupId or termGroupName but not both';
      }

      if (args.options.termGroupId) {
        if (!Utils.isValidGuid(args.options.termGroupId)) {
          return `${args.options.termGroupId} is not a valid GUID`;
        }
      }

      if (args.options.termSetId && args.options.termSetName) {
        return 'Specify termSetId or termSetName but not both';
      }

      if (args.options.termSetId) {
        if (!Utils.isValidGuid(args.options.termSetId)) {
          return `${args.options.termSetId} is not a valid GUID`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoTermGetCommand();