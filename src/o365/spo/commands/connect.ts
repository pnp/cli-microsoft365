import auth from '../SpoAuth';
import { ContextInfo } from '../spo';
import Auth from '../../../Auth';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import VerboseOption from '../../../VerboseOption';
import Command, {
  CommandAction,
  CommandCancel,
  CommandHelp,
  CommandValidate
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  url: string;
  options: VerboseOption;
}

const CONNECTION_SUCCEEDED: string = 'connection_succeeded';

class SpoConnectCommand extends Command {
  public get name(): string {
    return `${commands.CONNECT} <url>`;
  }

  public get description(): string {
    return 'Connects to a SharePoint Online site';
  }

  public get action(): CommandAction {
    return function (this: CommandInstance, args: CommandArgs, cb: () => void) {
      const chalk: any = vorpal.chalk;
      const verbose: boolean = args.options.verbose || false;

      appInsights.trackEvent({
        name: commands.CONNECT,
        properties: {
          verbose: verbose.toString()
        }
      });

      // disconnect before re-connecting
      if (verbose) {
        this.log(`
Disconnecting from SPO...
`);
      }
      auth.site.disconnect();

      this.log(`
Authenticating with SharePoint Online at ${args.url}...
`);

      const resource = Auth.getResourceFromUrl(args.url);

      auth
        .ensureAccessToken(resource, this, args.options.verbose)
        .then((accessToken: string): Promise<ContextInfo> => {
          auth.service.resource = resource;
          auth.site.url = args.url;
          this.log(chalk.green('DONE'));

          if (verbose) {
            this.log(`Checking if ${auth.site.url} is a tenant admin site...`);
          }
          if (auth.site.isTenantAdminSite()) {
            const requestDigestRequestOptions: any = {
              url: `${auth.site.url}/_api/contextinfo`,
              headers: {
                authorization: `Bearer ${accessToken}`,
                accept: 'application/json;odata=nometadata'
              },
              json: true
            };

            if (verbose) {
              this.log(`${auth.site.url} is a tenant admin site. Get tenant information...`);
              this.log('');
              this.log('Executing web request:');
              this.log(requestDigestRequestOptions);
              this.log('');
            }

            return request.post(requestDigestRequestOptions);
          }
          else {
            if (verbose) {
              this.log(`${auth.site.url} is not a tenant admin site`);
              this.log('');
            }

            auth.site.connected = true;
            this.log(`Successfully connected to ${args.url}`);
            cb();
            throw CONNECTION_SUCCEEDED;
          }
        })
        .then((res: ContextInfo): Promise<string> => {
          if (verbose) {
            this.log('Response:');
            this.log(res);
            this.log('');
          }

          const tenantInfoRequestOptions = {
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              authorization: `Bearer ${auth.site.accessToken}`,
              'X-RequestDigest': res.FormDigestValue,
              accept: 'application/json;odata=nometadata'
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
          };

          this.log('Retrieving tenant admin site information...');

          if (verbose) {
            this.log('Executing web request:');
            this.log(tenantInfoRequestOptions);
            this.log('');
          }

          return request.post(tenantInfoRequestOptions);
        })
        .then((res: string): void => {
          if (verbose) {
            this.log('Response:');
            this.log(res);
            this.log('');
          }

          const json: string[] = JSON.parse(res);

          auth.site.tenantId = (json[json.length - 1] as any)._ObjectIdentity_.replace('\n', '&#xA;');
          auth.site.connected = true;
          this.log(chalk.green('DONE'));
          this.log(`Successfully connected to ${args.url}
`);
          cb();
        }, (rej: Error | string): void => {
          if (rej instanceof Error) {
            if (verbose) {
              this.log('Error:');
              this.log(rej);
              this.log('');
            }

            this.log(chalk.red('Connecting to SharePoint Online failed'));
            this.log(`The following error occurred: ${rej.message}`);
            cb();
            return;
          }
          else {
            if (verbose) {
              this.log('Early exit of a promise chain');
            }
          }
        });
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.url.indexOf('https://') !== 0 ||
        args.url.indexOf('.sharepoint.com') === -1) {
        return `${args.url} is not a valid SharePoint Online URL`;
      }
      else {
        return true;
      }
    };
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (auth.interval) {
        clearInterval(auth.interval);
      }
    }
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.CONNECT).helpInformation());
      log(
        `  Arguments:
    
    url  absolute URL of the SharePoint Online site to connect to
        
  Remarks:

    Using the ${chalk.blue(commands.CONNECT)} command, you can connect to any SharePoint Online site.
    Depending on the command you want to use, you might be required to connect
    to a SharePoint Online tenant admin site (suffixed with ${chalk.grey('-admin')},
    eg. ${chalk.grey('https://contoso-admin.sharepoint.com')}) or a regular site.

    The ${chalk.blue(commands.CONNECT)} command uses device code OAuth flow with the standard
    Microsoft SharePoint Online Management Shell Azure AD application to connect
    to SharePoint Online.
    
    When connecting to a SharePoint site, the ${chalk.blue(commands.CONNECT)} command stores in memory
    the access token and the refresh token for the specified site. Both tokens are cleared from memory
    after exiting the CLI or by calling the ${chalk.blue(commands.DISCONNECT)} command.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.CONNECT} https://contoso-admin.sharepoint.com
      connects to a SharePoint Online tenant admin site

    ${chalk.grey(config.delimiter)} ${commands.CONNECT} --verbose https://contoso-admin.sharepoint.com
      connects to a SharePoint Online tenant admin site in verbose mode including
      detailed debug information in the console output
      
    ${chalk.grey(config.delimiter)} ${commands.CONNECT} https://contoso.sharepoint.com/sites/team
      connects to a regular SharePoint Online site
`);
    }
  }
}

module.exports = new SpoConnectCommand();