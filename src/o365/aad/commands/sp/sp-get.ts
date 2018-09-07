import auth from '../../AadAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../AadCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  displayName?: string;
}

class SpGetCommand extends AadCommand {
  public get name(): string {
    return commands.SP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific service principal';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = (!(!args.options.appId)).toString();
    telemetryProps.displayName = (!(!args.options.displayName)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about the service principal...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving service principal information...`);
        }

        const spMatchQuery: string = args.options.appId ?
          `appId eq '${encodeURIComponent(args.options.appId)}'` :
          `displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;

        const requestOptions: any = {
          url: `${auth.service.resource}/myorganization/servicePrincipals?api-version=1.6&$filter=${spMatchQuery}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { value: any[] }): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res, null, 2));
          cmd.log('');
        }

        if (res.value && res.value.length > 0) {
          cmd.log(res.value[0]);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]',
        description: 'ID of the application for which the service principal should be retrieved'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name of the application for which the service principal should be retrieved'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId && !args.options.displayName) {
        return 'Specify either appId or displayName';
      }

      if (args.options.appId) {
        if (!Utils.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }
      }

      if (args.options.appId && args.options.displayName) {
        return 'Specify either appId or displayName but not both';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SP_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to Azure Active Directory Graph,
      using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To get information about a service principal, you have to first log in to Azure Active Directory
    Graph using the ${chalk.blue(commands.LOGIN)} command.

    When looking up information about a service principal you should specify either its ${chalk.grey('appId')}
    or ${chalk.grey('displayName')} but not both. If you specify both values, the command will fail
    with an error.
   
  Examples:
  
    Return details about the service principal with appId ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}.
      ${chalk.grey(config.delimiter)} ${commands.SP_GET} --appId b2307a39-e878-458b-bc90-03bc578531d6

    Return details about the ${chalk.grey('Microsoft Graph')} service principal.
      ${chalk.grey(config.delimiter)} ${commands.SP_GET} --displayName "Microsoft Graph"

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD)
      https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new SpGetCommand();