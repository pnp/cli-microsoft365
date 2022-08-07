import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  cdnType: string;
  policy: string;
  value: string;
}

class SpoCdnPolicySetCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_POLICY_SET;
  }

  public get description(): string {
    return 'Sets CDN policy value for the current SharePoint Online tenant';
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
        cdnType: args.options.cdnType || 'Public',
        policy: args.options.policy
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --cdnType [cdnType]',
        autocomplete: ['Public', 'Private']
      },
      {
        option: '-p, --policy <policy>',
        autocomplete: ['IncludeFileExtensions', 'ExcludeRestrictedSiteClassifications']
      },
      {
        option: '-v, --value <value>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.cdnType) {
          if (args.options.cdnType !== 'Public' &&
            args.options.cdnType !== 'Private') {
            return `${args.options.cdnType} is not a valid CDN type. Allowed values are Public|Private`;
          }
        }

        if (!args.options.policy ||
          (args.options.policy !== 'IncludeFileExtensions' &&
            args.options.policy !== 'ExcludeRestrictedSiteClassifications')) {
          return `${args.options.policy} is not a valid CDN policy. Allowed values are IncludeFileExtensions|ExcludeRestrictedSiteClassifications`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.cdnType || 'Public';
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
          logger.logToStderr(`Configuring policy on the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
        }

        let policyId: number = -1;
        switch (args.options.policy) {
          case "IncludeFileExtensions":
            policyId = 0;
            break;
          case "ExcludeRestrictedSiteClassifications":
            policyId = 1;
            break;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Enum">${policyId}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.value)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${tenantId}" /></ObjectPaths></Request>`
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

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoCdnPolicySetCommand();