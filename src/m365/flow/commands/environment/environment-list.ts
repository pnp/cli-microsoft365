import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class FlowEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Retrieving list of Microsoft Flow environments...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ name: string, properties: { displayName: string } }] }>(requestOptions)
      .then((res: { value: [{ name: string, properties: { displayName: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            logger.log(res.value);
          }
          else {
            logger.log(res.value.map(e => {
              return {
                name: e.name,
                displayName: e.properties.displayName
              };
            }));
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new FlowEnvironmentListCommand();