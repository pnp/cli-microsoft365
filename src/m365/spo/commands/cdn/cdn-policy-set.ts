import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    telemetryProps.policy = args.options.policy;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    this
      .getTenantId(logger, this.debug)
      .then((_tenantId: string): Promise<string> => {
        tenantId = _tenantId;
        return this.getSpoAdminUrl(logger, this.debug);
      })
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
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
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Enum">${policyId}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.value)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${tenantId}" /></ObjectPaths></Request>`
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
          if (this.verbose) {
            logger.logToStderr(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --type [type]',
        description: 'Type of CDN to manage. Public|Private. Default Public',
        autocomplete: ['Public', 'Private']
      },
      {
        option: '-p, --policy <policy>',
        description: 'CDN policy to configure. IncludeFileExtensions|ExcludeRestrictedSiteClassifications',
        autocomplete: ['IncludeFileExtensions', 'ExcludeRestrictedSiteClassifications']
      },
      {
        option: '-v, --value <value>',
        description: 'Value for the policy to configure'
      }
    ];

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

    if (!args.options.policy ||
      (args.options.policy !== 'IncludeFileExtensions' &&
        args.options.policy !== 'ExcludeRestrictedSiteClassifications')) {
      return `${args.options.policy} is not a valid CDN policy. Allowed values are IncludeFileExtensions|ExcludeRestrictedSiteClassifications`;
    }

    return true;
  }
}

module.exports = new SpoCdnPolicySetCommand();