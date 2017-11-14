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

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  type: string;
  enabled: string;
}

class SpoTenantCdnSetCommand extends Command {
  public get name(): string {
    return commands.TENANT_CDN_SET;
  }

  public get description(): string {
    return 'Enable or disable the specified Office 365 CDN';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;
      const cdnTypeString: string = args.options.type || 'Public';
      const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
      const enabled: boolean = args.options.enabled === 'true';

      appInsights.trackEvent({
        name: commands.TENANT_CDN_SET,
        properties: {
          cdnType: cdnTypeString,
          enable: enabled.toString(),
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

      auth
        .ensureAccessToken(auth.service.resource, this, verbose)
        .then((accessToken: string): Promise<ContextInfo> => {
          if (verbose) {
            this.log(`Retrieved access token ${accessToken}. Loading CDN settings for the ${auth.site.url} tenant...`);
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_api/contextinfo`,
            headers: {
              authorization: `Bearer ${accessToken}`,
              accept: 'application/json;odata=nometadata'
            },
            json: true
          }

          if (verbose) {
            this.log('Executing web request...');
            this.log(requestOptions);
            this.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (verbose) {
            this.log('Response:');
            this.log(res);
            this.log('');
          }

          this.log(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);

          const requestOptions: any = {
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            this.log('Response:');
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

      const enabled: string | undefined = args.options.enabled ? args.options.enabled.toLowerCase() : undefined;
      if (enabled !== 'true' &&
        enabled !== 'false') {
        return `${args.options.enabled} is not a valid boolean value. Allowed values are true|false`;
      }

      return true;
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --enabled <enabled>',
        description: 'Set to true to enable CDN or to false to disable it. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '-t, --type [type]',
        description: 'Type of CDN to manage. Public|Private. Default Public',
        autocomplete: ['Public', 'Private']
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

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.TENANT_CDN_SET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To enable or disable an Office 365 CDN, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Using the ${chalk.blue('-t, --type')} option you can choose whether you want to manage the settings of
    the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

    Using the ${chalk.blue('-e, --enabled')} option you can specify whether the given CDN type should be
    enabled or disabled. Use ${chalk.grey('true')} to enable the specified CDN and ${chalk.grey('false')} to
    disable it.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.TENANT_CDN_SET} -t Public -e true
      enables the Office 365 Public CDN on the current tenant

    ${chalk.grey(config.delimiter)} ${commands.TENANT_CDN_SET} -t Public -e false
      disables the Office 365 Public CDN on the current tenant

  More information:

    General availability of Office 365 CDN
      https://dev.office.com/blogs/general-availability-of-office-365-cdn
`);
    };
  }
}

module.exports = new SpoTenantCdnSetCommand();