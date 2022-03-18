import { Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
}

class SpoCdnOriginListCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_ORIGIN_LIST;
  }

  public get description(): string {
    return 'List CDN origins settings for the current SharePoint Online tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    spo
      .getTenantId(logger, this.debug)
      .then((_tenantId: string): Promise<string> => {
        tenantId = _tenantId;
        return spo.getSpoAdminUrl(logger, this.debug);
      })
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Retrieving configured origins for ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnOrigins" Id="22" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="18" Name="${tenantId}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: string[] = json[json.length - 1];
          logger.log(result);
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-t, --type [type]',
      autocomplete: ['Public', 'Private']
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.type) {
      if (args.options.type !== 'Public' &&
        args.options.type !== 'Private') {
        return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
      }
    }

    return true;
  }
}

module.exports = new SpoCdnOriginListCommand();