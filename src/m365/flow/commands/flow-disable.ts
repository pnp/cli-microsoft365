import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import { formatting } from '../../../utils/formatting';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environmentName: string;
  asAdmin: boolean;
}

class FlowDisableCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.DISABLE;
  }

  public get description(): string {
    return 'Disables specified Microsoft Flow';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Disables Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/stop?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowDisableCommand();