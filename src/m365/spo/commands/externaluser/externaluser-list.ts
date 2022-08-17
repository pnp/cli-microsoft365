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
import { GetExternalUsersResults } from './GetExternalUsersResults';

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
    return commands.EXTERNALUSER_LIST;
  }

  public get description(): string {
    return 'Lists external users in the tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        filter: (!(!args.options.filter)).toString(),
        pageSize: (!(!args.options.pageSize)).toString(),
        position: (!(!args.options.position)).toString(),
        sortOrder: (!(!args.options.sortOrder)).toString(),
        siteUrl: (!(!args.options.siteUrl)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --filter [filter]'
      },
      {
        option: '-p, --pageSize [pageSize]'
      },
      {
        option: '-i, --position [position]'
      },
      {
        option: '-s, --sortOrder [sortOrder]',
        autocomplete: ['asc', 'desc']
      },
      {
        option: '-u, --siteUrl [siteUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
          return validation.isValidSharePointUrl(args.options.siteUrl);
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Retrieving information about external users...`);
        }

        const position: number = parseInt(args.options.position || '0');
        const pageSize: number = parseInt(args.options.pageSize || '10');
        const sortOrder: number = args.options.sortOrder === 'desc' ? 1 : 0;

        let payload: string = '';

        if (args.options.siteUrl) {
          payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
        }
        else {
          payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: payload
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
            logger.log(results.ExternalUserCollection._Child_Items_.map(e => {
              delete e._ObjectType_;
              const dateChunks: number[] = (e.WhenCreated as string)
                .replace('/Date(', '')
                .replace(')/', '')
                .split(',')
                .map(c => {
                  return parseInt(c);
                });
              e.WhenCreated = new Date(dateChunks[0], dateChunks[1], dateChunks[2], dateChunks[3], dateChunks[4], dateChunks[5], dateChunks[6]);
              return e;
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoExternalUserListCommand();