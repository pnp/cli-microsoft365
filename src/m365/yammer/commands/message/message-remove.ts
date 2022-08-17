import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  confirm?: boolean;
}

class YammerMessageRemoveCommand extends YammerCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Yammer message';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeMessage: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1/messages/${args.options.id}.json`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .delete(requestOptions)
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeMessage();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Yammer message ${args.options.id}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeMessage();
        }
      });
    }
  }
}

module.exports = new YammerMessageRemoveCommand();
