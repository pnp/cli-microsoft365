import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        cdnType: args.options.type || 'Public'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: ['Public', 'Private']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type) {
          if (args.options.type !== 'Public' &&
            args.options.type !== 'Private') {
            return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    try {
      tenantId = await spo.getTenantId(logger, this.debug);
      spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        logger.logToStderr(`Retrieving configured origins for ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnOrigins" Id="22" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="18" Name="${tenantId}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const result: string[] = json[json.length - 1];
        logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoCdnOriginListCommand();