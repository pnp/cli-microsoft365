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
import { TermGroup } from './TermGroup';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class SpoTermGroupGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_GROUP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified taxonomy term group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
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
          cmd.log(`Retrieving taxonomy term groups...`);
        }

        const query: string = args.options.id ? `<Method Id="32" ParentId="30" Name="GetById"><Parameters><Parameter Type="Guid">{${Utils.escapeXml(args.options.id)}}</Parameter></Parameters></Method>` : `<Method Id="32" ParentId="30" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" />${query}</ObjectPaths></Request>`
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

        const termGroup: TermGroup = json[json.length - 1];
        delete termGroup._ObjectIdentity_;
        delete termGroup._ObjectType_;
        termGroup.CreatedDate = new Date(Number(termGroup.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        termGroup.Id = termGroup.Id.replace('/Guid(', '').replace(')/', '');
        termGroup.LastModifiedDate = new Date(Number(termGroup.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        cmd.log(termGroup);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the term group to retrieve. Specify name or id but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the term group to retrieve. Specify name or id but not both'
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

      return true;
    };
  }
}

module.exports = new SpoTermGroupGetCommand();