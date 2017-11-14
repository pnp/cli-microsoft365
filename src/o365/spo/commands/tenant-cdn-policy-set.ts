import auth from '../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../spo';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import VerboseOption from '../../../VerboseOption';
import Command, {
  CommandAction,
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../Command';
import appInsights from '../../../appInsights';
import Utils from '../../../Utils';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  type: string;
  policy: string;
  value: string;
}

class SpoTenantCdnPolicySetCommand extends Command {
  public get name(): string {
    return commands.TENANT_CDN_POLICY_SET;
  }

  public get description(): string {
    return 'Sets CDN policy value for the current SharePoint Online tenant';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;
      const cdnTypeString: string = args.options.type || 'Public';
      const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;

      appInsights.trackEvent({
        name: commands.TENANT_CDN_POLICY_SET,
        properties: {
          cdnType: cdnTypeString,
          policy: args.options.policy,
          verbose: verbose.toString()
        }
      });

      if (!auth.site.connected) {
        this.log('Connect to a SharePoint Online tenant admin site first');
        cb();
        return;
      }

      if (!auth.site.isTenantAdminSite()) {
        this.log(`${auth.site.url} is not a tenant admin site. Connect to your tenant admin site and try again`);
        cb();
        return;
      }

      if (verbose) {
        this.log(`Retrieving access token for ${auth.service.resource}...`);
      }

      auth
        .ensureAccessToken(auth.service.resource, this, verbose)
        .then((accessToken: string): Promise<ContextInfo> => {
          if (verbose) {
            this.log('Response:');
            this.log(accessToken);
            this.log('');
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_api/contextinfo`,
            headers: {
              authorization: `Bearer ${accessToken}`,
              accept: 'application/json;odata=nometadata'
            },
            json: true
          };

          if (verbose) {
            this.log('Executing web request...');
            this.log(requestOptions);
            this.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (verbose) {
            this.log('Response:')
            this.log(res);
            this.log('');
          }

          this.log(`Configuring policy on the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);

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
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              authorization: `Bearer ${auth.site.accessToken}`,
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Enum">${policyId}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.value)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${auth.site.tenantId}" /></ObjectPaths></Request>`
          };

          if (verbose) {
            this.log('Executing web request...');
            this.log(requestOptions);
            this.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          if (verbose) {
            this.log('Response:')
            this.log(res);
            this.log('');
          }

          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            this.log(vorpal.chalk.red(`Error: ${response.ErrorInfo.ErrorMessage}`));
          }
          else {
            this.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => {
          this.log(vorpal.chalk.red(`Error: ${err}`));
          cb();
        });
    };
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

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.TENANT_CDN_POLICY_SET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To set the policy of an Office 365 CDN, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Using the ${chalk.blue('-t, --type')} option you can choose whether you want to manage the settings of
    the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.TENANT_CDN_POLICY_SET} -t Public -p IncludeFileExtensions -v CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON
      sets the list of extensions supported by the Public CDN

  More information:

    General availability of Office 365 CDN
      https://dev.office.com/blogs/general-availability-of-office-365-cdn
`);
    };
  }
}

module.exports = new SpoTenantCdnPolicySetCommand();