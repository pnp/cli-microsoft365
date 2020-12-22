import { Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  flow: string;
}

class FlowRunListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_RUN_LIST;
  }

  public get description(): string {
    return 'Lists runs of the specified Microsoft Flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'startTime', 'status'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of runs for Microsoft Flow ${args.options.flow}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/runs?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ name: string; startTime: string; status: string; properties: { startTime: string, status: string } }] }>(requestOptions)
      .then((res: { value: [{ name: string, startTime: string; status: string; properties: { startTime: string, status: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          res.value.forEach(r => {
            r.startTime = r.properties.startTime;
            r.status = r.properties.status;
          });

          logger.log(res.value);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No runs found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-f, --flow <flow>',
        description: 'The name of the Microsoft Flow to retrieve the runs for'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment to which the flow belongs'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowRunListCommand();