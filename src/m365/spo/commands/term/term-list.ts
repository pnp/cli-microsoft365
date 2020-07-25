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
import { TermCollection } from './TermCollection';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  termGroupId?: string;
  termGroupName?: string;
  termSetId?: string;
  termSetName?: string;
}

class SpoTermListCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_LIST;
  }

  public get description(): string {
    return 'Lists taxonomy terms from the given term set';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
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
          cmd.log(`Retrieving taxonomy term sets...`);
        }

        const termGroupQuery: string = args.options.termGroupId ? `<Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
        const termSetQuery: string = args.options.termSetId ? `<Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" />${termGroupQuery}<Property Id="80" ParentId="77" Name="TermSets" />${termSetQuery}<Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`
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

        const result: TermCollection = json[json.length - 1];
        if (result._Child_Items_ && result._Child_Items_.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(result._Child_Items_.map(t => {
              t.CreatedDate = new Date(Number(t.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
              t.Id = t.Id.replace('/Guid(', '').replace(')/', '');
              t.LastModifiedDate = new Date(Number(t.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
              return t;
            }));
          }
          else {
            cmd.log(result._Child_Items_.map(t => {
              return {
                Id: t.Id.replace('/Guid(', '').replace(')/', ''),
                Name: t.Name
              };
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--termGroupId [termGroupId]',
        description: 'ID of the term group where the term set is located. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termGroupName [termGroupName]',
        description: 'Name of the term group where the term set is located. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termSetId [termSetId]',
        description: 'ID of the term set for which to retrieve terms. Specify termSetId or termSetName but not both'
      },
      {
        option: '--termSetName [termSetName]',
        description: 'Name of the term set for which to retrieve terms. Specify termSetId or termSetName but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.termGroupId && !args.options.termGroupName) {
        return 'Specify either termGroupId or termGroupName';
      }

      if (args.options.termGroupId && args.options.termGroupName) {
        return 'Specify either termGroupId or termGroupName but not both';
      }

      if (args.options.termGroupId) {
        if (!Utils.isValidGuid(args.options.termGroupId)) {
          return `${args.options.termGroupId} is not a valid GUID`;
        }
      }

      if (!args.options.termSetId && !args.options.termSetName) {
        return 'Specify either termSetId or termSetName';
      }

      if (args.options.termSetId && args.options.termSetName) {
        return 'Specify either termSetId or termSetName but not both';
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

module.exports = new SpoTermListCommand();