import * as chalk from 'chalk';
import { Cli,Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  flow: string;
  name: string;
}

class FlowRunCancelCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_RUN_CANCEL;
  }

  public get description(): string {
    return 'Cancelling a specific run of the specified Microsoft Flow';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Cancelling run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
    }
  
    const cancelFlow: () => void = (): void => {
    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/runs/${encodeURIComponent(args.options.name)}/cancel?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((rawRes: any): void => {
        // handle 204 and throw error message to cmd when invalid flow id is passed
        // https://github.com/pnp/cli-microsoft365/issues/1063#issuecomment-537218957
        debugger;
        if (rawRes.statusCode === 204) {
          logger.log(chalk.red(`Error: Resource '${args.options.name}' does not exist in environment '${args.options.environment}'`));
          cb();
        }
        else {
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }
          cb();
        }
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      cancelFlow();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to cancel the flow with runId ${args.options.name}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          cancelFlow();
        }
      });
    }
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
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming cancelling the Flow'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.flow)) {
      return `${args.options.flow} is not a valid GUID`;
    }

    //Is this needed?
    if (!Utils.isValidFlowRunId(args.options.name)) {
      return `${args.options.name} is not a valid RUN ID`;
    }

    return true;
  }
}

module.exports = new FlowRunCancelCommand();