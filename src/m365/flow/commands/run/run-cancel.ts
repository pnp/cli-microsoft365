import { Cli, Logger } from '../../../../cli';
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
    return 'Cancels a specific run of the specified Microsoft Flow';
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
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      cancelFlow();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to cancel the flow run ${args.options.name}?`
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
        option: '-n, --name <name>'
      },
      {
        option: '-f, --flow <flow>'
      },
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.flow)) {
      return `${args.options.flow} is not a valid GUID`;
    }
    
    return true;
  }
}

module.exports = new FlowRunCancelCommand();