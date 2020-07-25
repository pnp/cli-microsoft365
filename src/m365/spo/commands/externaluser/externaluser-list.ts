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
import { GetExternalUsersResults } from './GetExternalUsersResults';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filter?: string;
  pageSize?: string;
  position?: string;
  sortOrder?: string;
  siteUrl?: string;
}

class SpoExternalUserListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.EXTERNALUSER_LIST}`;
  }

  public get description(): string {
    return 'Lists external users in the tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.filter = (!(!args.options.filter)).toString();
    telemetryProps.pageSize = (!(!args.options.pageSize)).toString();
    telemetryProps.position = (!(!args.options.position)).toString();
    telemetryProps.sortOrder = (!(!args.options.sortOrder)).toString();
    telemetryProps.siteUrl = (!(!args.options.siteUrl)).toString();
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
          cmd.log(`Retrieving information about external users...`);
        }

        const position: number = parseInt(args.options.position || '0');
        const pageSize: number = parseInt(args.options.pageSize || '10');
        const sortOrder: number = args.options.sortOrder === 'desc' ? 1 : 0;

        let payload: string = '';

        if (args.options.siteUrl) {
          payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.siteUrl)}</Parameter><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
        }
        else {
          payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: payload
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
        else {
          const results: GetExternalUsersResults = json.pop();

          if (results.TotalUserCount > 0) {
            cmd.log(results.ExternalUserCollection._Child_Items_.map(e => {
              delete e._ObjectType_;
              const dateChunks: number[] = (e.WhenCreated as string)
                .replace('/Date(', '')
                .replace(')/', '')
                .split(',')
                .map(c => {
                  return parseInt(c);
                })
              e.WhenCreated = new Date(dateChunks[0], dateChunks[1], dateChunks[2], dateChunks[3], dateChunks[4], dateChunks[5], dateChunks[6]);
              return e;
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-f, --filter [filter]',
        description: 'Limits the results to only those users whose first name, last name or email address begins with the text in the string, using a case-insensitive comparison'
      },
      {
        option: '-p, --pageSize [pageSize]',
        description: 'Specifies the maximum number of users to be returned in the collection. The value must be less than or equal to 50'
      },
      {
        option: '-i, --position [position]',
        description: 'Use to specify the zero-based index of the position in the sorted collection of the first result to be returned'
      },
      {
        option: '-s, --sortOrder [sortOrder]',
        description: 'Specifies the sort results in Ascending or Descending order on the SPOUser.Email property should occur. Allowed values asc|desc. Default asc',
        autocomplete: ['asc', 'desc']
      },
      {
        option: '-u, --siteUrl [siteUrl]',
        description: 'Specifies the site to retrieve external users for. If no site is specified, the external users for all sites are returned'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.pageSize) {
        const pageSize: number = parseInt(args.options.pageSize);
        if (isNaN(pageSize)) {
          return `${args.options.pageSize} is not a valid number`;
        }

        if (pageSize < 1 || pageSize > 50) {
          return 'pageSize must be between 1 and 50';
        }
      }

      if (args.options.position) {
        const position: number = parseInt(args.options.position);
        if (isNaN(position)) {
          return `${args.options.position} is not a valid number`;
        }

        if (position < 0) {
          return 'position must be greater than or 0';
        }
      }

      if (args.options.sortOrder &&
        args.options.sortOrder !== 'asc' &&
        args.options.sortOrder !== 'desc') {
        return `${args.options.sortOrder} is not a valid sortOrder value. Allowed values asc|desc`;
      }

      if (args.options.siteUrl) {
        return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
      }

      return true;
    };
  }
}

module.exports = new SpoExternalUserListCommand();