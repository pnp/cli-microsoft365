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
import * as os from 'os';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  name: string;
  asAdmin: boolean;
}

class AzmgmtFlowGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information about Microsoft Flow ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
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

        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          const summary: any = {
            name: res.name,
            displayName: res.properties.displayName,
            description: res.properties.definitionSummary.description || '',
            triggers: Object.keys(res.properties.definition.triggers).join(', '),
            actions: Object.keys(res.properties.definition.actions).join(', ')
          };
          cmd.log(summary);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the Microsoft Flow to get information about'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to retrieve available Flows'
      },
      {
        option: '--asAdmin',
        description: 'Set, to retrieve the Flow as admin'
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

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Azure Management Service,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To get information about the specified Microsoft Flow, you have to first
    connect to the Azure Management Service using the ${chalk.blue(commands.CONNECT)} command.

    By default, the command will try to retrieve Microsoft Flows you own.
    If you want to retrieve Flow owned by another user, use the ${chalk.blue('asAdmin')}
    flag.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the name you specified doesn't exist, you will
    get the ${chalk.grey(`The caller with object id \'abc\' does not have permission${os.EOL}` +
    '    for connection \'xyz\' under Api \'shared_logicflows\'.')} error.
    If you try to retrieve a non-existing flow as admin, you will get the
    ${chalk.grey('Could not find flow \'xyz\'.')} error.
   
  Examples:
  
    Get information about the specified Microsoft Flow owned by the currently
    signed-in user
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d

    Get information about the specified Microsoft Flow owned by another user
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
`);
  }
}

module.exports = new AzmgmtFlowGetCommand();