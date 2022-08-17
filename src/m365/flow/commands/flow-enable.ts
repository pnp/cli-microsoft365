import { Logger } from '../../../cli';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environment: string;
  asAdmin: boolean;
}

class FlowEnableCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.ENABLE;
  }

  public get description(): string {
    return 'Enables specified Microsoft Flow';
  }

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--asAdmin'
      }      
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Enables Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}/start?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new FlowEnableCommand();