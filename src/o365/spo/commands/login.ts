import auth from '../SpoAuth';
import { ContextInfo } from '../spo';
import { Auth, AuthType } from '../../../Auth';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import Command, {
  CommandCancel,
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../Command';
import SpoCommand from '../SpoCommand';
import Utils from '../../../Utils';
import appInsights from '../../../appInsights';
import * as fs from 'fs';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  url: string;
  options: Options;
}

interface Options extends GlobalOptions {
  authType?: string;
  userName?: string;
  password?: string;
  certificateFile?: string;
  thumbprint?: string;
}

class SpoLoginCommand extends Command {
  public get name(): string {
    return `${commands.LOGIN} <url>`;
  }

  public get description(): string {
    return 'Log in to SharePoint Online';
  }

  public alias(): string[] | undefined {
    return [commands.CONNECT];
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const chalk: any = vorpal.chalk;

    this.showDeprecationWarning(cmd, commands.CONNECT, commands.LOGIN);

    appInsights.trackEvent({
      name: this.getUsedCommandName(cmd)
    });

    // disconnect before re-connecting
    if (this.debug) {
      cmd.log(`Logging out from SPO...`);
    }

    const logout: () => void = (): void => {
      auth.site.logout();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
    }

    const login: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Logging in to SharePoint Online at ${args.url}...`);
      }

      const resource = Auth.getResourceFromUrl(args.url);
      auth.site.url = args.url;

      switch (args.options.authType) {
        case 'password':
          auth.service.authType = AuthType.Password;
          auth.service.userName = args.options.userName;
          auth.service.password = args.options.password;
          break;
        case 'certificate':
          auth.service.authType = AuthType.Certificate;
          auth.service.certificate = fs.readFileSync(args.options.certificateFile as string, 'utf8');
          auth.service.thumbprint = args.options.thumbprint;
          break;
      }

      if (auth.site.isTenantAdminSite()) {
        auth
          .ensureAccessToken(resource, cmd, this.debug)
          .then((accessToken: string): request.RequestPromise => {
            auth.service.resource = resource;
            auth.site.url = args.url;
            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }

            const requestDigestRequestOptions: any = {
              url: `${auth.site.url}/_api/contextinfo`,
              headers: Utils.getRequestHeaders({
                authorization: `Bearer ${accessToken}`,
                accept: 'application/json;odata=nometadata'
              }),
              json: true
            };

            if (this.debug) {
              cmd.log(`${auth.site.url} is a tenant admin site. Get tenant information...`);
              cmd.log('');
              cmd.log('Executing web request:');
              cmd.log(requestDigestRequestOptions);
              cmd.log('');
            }

            return request.post(requestDigestRequestOptions);
          })
          .then((res: ContextInfo): request.RequestPromise => {
            if (this.debug) {
              cmd.log('Response:');
              cmd.log(res);
              cmd.log('');
            }

            const tenantInfoRequestOptions = {
              url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
              headers: Utils.getRequestHeaders({
                authorization: `Bearer ${auth.site.accessToken}`,
                'X-RequestDigest': res.FormDigestValue,
                accept: 'application/json;odata=nometadata'
              }),
              body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };

            if (this.verbose) {
              cmd.log('Retrieving tenant admin site information...');
            }

            if (this.debug) {
              cmd.log('Executing web request:');
              cmd.log(tenantInfoRequestOptions);
              cmd.log('');
            }

            return request.post(tenantInfoRequestOptions);
          })
          .then((res: string): Promise<void> => {
            if (this.debug) {
              cmd.log('Response:');
              cmd.log(res);
              cmd.log('');
            }

            const json: string[] = JSON.parse(res);

            auth.site.tenantId = (json[json.length - 1] as any)._ObjectIdentity_.replace('\n', '&#xA;');
            auth.site.connected = true;
            return auth.storeSiteConnectionInfo();
          })
          .then((): void => {
            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
              cmd.log(`Successfully logged in to ${args.url}`);
            }
            cb();
          }, (rej: string): void => {
            if (this.debug) {
              cmd.log('Error:');
              cmd.log(rej);
              cmd.log('');
            }

            if (rej !== 'Polling_Request_Cancelled') {
              cb(new CommandError(rej));
              return;
            }
            cb();
            return;
          });
      }
      else {
        auth
          .ensureAccessToken(resource, cmd, this.debug)
          .then((accessToken: string): Promise<void> => {
            auth.service.resource = resource;
            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }

            auth.site.connected = true;
            return auth.storeSiteConnectionInfo();
          })
          .then((): void => {
            cb();
          }, (rej: string): void => {
            if (this.debug) {
              cmd.log('Error:');
              cmd.log(rej);
              cmd.log('');
            }

            if (rej !== 'Polling_Request_Cancelled') {
              cb(new CommandError(rej));
              return;
            }
            cb();
          });
      }
    }

    auth
      .clearSiteConnectionInfo()
      .then((): void => {
        logout();
        login();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        logout();
        login();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --authType [authType]',
        description: 'The type of authentication to use. Allowed values certificate|deviceCode|password. Default deviceCode',
        autocomplete: ['certificate', 'deviceCode', 'password']
      },
      {
        option: '-u, --userName [userName]',
        description: 'Name of the user to authenticate. Required when authType is set to password'
      },
      {
        option: '-p, --password [password]',
        description: 'Password for the user. Required when authType is set to password'
      },
      {
        option: '-c, --certificateFile [certificateFile]',
        description: 'Path to the file with certificate private key. Required when authType is set to certificate'
      },
      {
        option: '--thumbprint [thumbprint]',
        description: 'Certificate thumbprint. Required when authType is set to certificate'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.authType === 'password') {
        if (!args.options.userName) {
          return 'Required option userName missing';
        }

        if (!args.options.password) {
          return 'Required option password missing';
        }
      }

      if (args.options.authType === 'certificate') {
        if (!args.options.certificateFile) {
          return 'Required option certificateFile missing';
        }

        if (!fs.existsSync(args.options.certificateFile)) {
          return `File '${args.options.certificateFile}' does not exist`;
        }

        if (!args.options.thumbprint) {
          return 'Required option thumbprint missing';
        }
      }

      return SpoCommand.isValidSharePointUrl(args.url);
    };
  }

  public cancel(): CommandCancel {
    return (): void => {
      auth.cancel();
    }
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.LOGIN).helpInformation());
    log(
      `  Arguments:
    
    url  absolute URL of the SharePoint Online site to log in to
        
  Remarks:

    Using the ${chalk.blue(commands.LOGIN)} command, you can log in to any SharePoint Online
    site. Depending on the command you want to use, you might be required to
    log in to a SharePoint Online tenant admin site (suffixed with ${chalk.grey('-admin')},
    eg. ${chalk.grey('https://contoso-admin.sharepoint.com')}) or a regular site.

    By default, the ${chalk.blue(commands.LOGIN)} command uses device code OAuth flow
    to log in to SharePoint Online. Alternatively, you can
    authenticate using a user name and password or certificate, which are
    convenient for CI/CD scenarios, but which come with their own limitations.
    See the Office 365 CLI manual for more information.
    
    When logging in to a SharePoint site, the ${chalk.blue(commands.LOGIN)} command
    stores in memory the access token and the refresh token for the specified
    site. Both tokens are cleared from memory after exiting the CLI or by
    calling the ${chalk.blue(commands.LOGOUT)} command.

    When logging in to SharePoint Online using the user name and
    password, next to the access and refresh token, the Office 365 CLI will
    store the user credentials so that it can automatically re-authenticate if
    necessary. Similarly to the tokens, the credentials are removed by
    re-authenticating using the device code or by calling the ${chalk.blue(commands.LOGOUT)}
    command.

    When logging in to SharePoint Online using a certificate, the Office 365 CLI
    will store the contents of the certificate so that it can automatically
    re-authenticate if necessary. The contents of the certificate are removed
    by re-authenticating using the device code or by calling the ${chalk.blue(commands.LOGOUT)}
    command.

    To log in to SharePoint Online using a certificate, you will typically
    create a custom Azure AD application. To use this application
    with the Office 365 CLI, you will set the ${chalk.grey('OFFICE365CLI_AADAADAPPID')}
    environment variable to the application's ID and the ${chalk.grey('OFFICE365CLI_TENANT')}
    environment variable to the ID of the Azure AD tenant, where you created
    the Azure AD application.

  Examples:
  
    Log in to a SharePoint Online tenant admin site using the device code
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} https://contoso-admin.sharepoint.com

    Log in to a SharePoint Online tenant admin site using the device code in
    debug mode including detailed debug information in the console output
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} --debug https://contoso-admin.sharepoint.com
      
    Log in to a regular SharePoint Online site using the device code
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} https://contoso.sharepoint.com/sites/team

    Log in to a SharePoint Online tenant admin site using a user name and password
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} https://contoso-admin.sharepoint.com --authType password --userName user@contoso.com --password pass@word1

    Log in to a SharePoint Online tenant admin site using a certificate
      ${chalk.grey(config.delimiter)} ${commands.LOGIN} https://contoso-admin.sharepoint.com --authType certificate --certificateFile /Users/user/dev/localhost.pfx --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1
`);
  }
}

module.exports = new SpoLoginCommand();