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
  flow: string;
  name: string;
}

class AzmgmtFlowRunGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_RUN_GET;
  }

  public get description(): string {
    return 'Gets information about a specific run of the specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information about run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/runs/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
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
            startTime: res.properties.startTime,
            endTime: res.properties.endTime || '',
            status: res.properties.status,
            triggerName: res.properties.trigger.name
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
        description: 'The name of the run to get information about'
      },
      {
        option: '-f, --flow <flow>',
        description: 'The name of the Microsoft Flow for which to retrieve information'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment where the Flow is located'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.flow) {
        return 'Required option flow missing';
      }

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      if (!args.options.name) {
        return 'Required option name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_RUN_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Azure Management Service,
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To get information about the specified run, you have to first log in to
    the Azure Management Service using the ${chalk.blue(commands.LOGIN)} command.
 
    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the name you specified doesn't exist, you will
    get the ${chalk.grey(`The caller with object id \'abc\' does not have permission${os.EOL}` +
        '    for connection \'xyz\' under Api \'shared_logicflows\'.')} error.

    If the run with the name you specified doesn't exist, you will
    get the ${chalk.grey(`The provided workflow run name is not valid.`)} error.
   
  Examples:
  
    Get information about the given run of the specified Microsoft Flow
      ${chalk.grey(config.delimiter)} ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flow 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62
`);
  }
}

module.exports = new AzmgmtFlowRunGetCommand();