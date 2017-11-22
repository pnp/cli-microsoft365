import auth from '../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../spo';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import VerboseOption from '../../../VerboseOption';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../Command';
import SpoCommand from '../SpoCommand';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  type: string;
}

class SpoCdnGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_GET;
  }

  public get description(): string {
    return 'View current status of the specified Office 365 CDN';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.verbose)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (this.verbose) {
          cmd.log(`Retrieved access token ${accessToken}. Loading CDN settings for the ${auth.site.url} tenant...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/contextinfo`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        if (this.verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Retrieving status of ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${auth.site.tenantId}" /></ObjectPaths></Request>`
        };

        if (this.verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cmd.log(vorpal.chalk.red(`Error: ${response.ErrorInfo.ErrorMessage}`));
        }
        else {
          const result: boolean = json[json.length - 1];
          cmd.log(`${(cdnType === 0 ? 'Public' : 'Private')} CDN at ${auth.site.url} is ${(result === true ? 'enabled' : 'disabled')}`);
        }
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-t, --type [type]',
      description: 'Type of CDN to manage. Public|Private. Default Public',
      autocomplete: ['Public', 'Private']
    }];

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

      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.CDN_GET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To view the status of an Office 365 CDN, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Using the ${chalk.blue('-t, --type')} option you can choose whether you want to manage the settings of
    the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.CDN_GET}
      shows if the Public CDN is currently enabled or not

    ${chalk.grey(config.delimiter)} ${commands.CDN_GET} -t Private
      shows if the Private CDN is currently enabled or not

  More information:

    General availability of Office 365 CDN
      https://dev.office.com/blogs/general-availability-of-office-365-cdn
`);
    };
  }
}

module.exports = new SpoCdnGetCommand();