import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import YammerCommand from '../../../base/YammerCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  force?: boolean;
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '-f, --force'
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeMessage: () => Promise<void> = async (): Promise<void> => {
      try {
        const requestOptions: any = {
          url: `${this.resource}/v1/messages/${args.options.id}.json`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeMessage();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Yammer message ${args.options.id}?`
      });

      if (result.continue) {
        await removeMessage();
      }
    }
  }
}

export default new YammerMessageRemoveCommand();
