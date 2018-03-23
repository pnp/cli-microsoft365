import auth from '../../AzmgmtAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import AzmgmtCommand from '../../AzmgmtCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class AzmgmtFlowEnvironmentGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow environment';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information about Microsoft Flow environment ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json'
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
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the environment to get information about'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required option name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_ENVIRONMENT_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Azure Management Service,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To get information about the specified Microsoft Flow environment, you have
    to first connect to the Azure Management Service using the
    ${chalk.blue(commands.CONNECT)} command.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.
   
  Examples:
  
    Get information about the Microsoft Flow environment named
    ${chalk.grey('Default-d87a7535-dd31-4437-bfe1-95340acd55c5')}
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
`);
  }
}

module.exports = new AzmgmtFlowEnvironmentGetCommand();