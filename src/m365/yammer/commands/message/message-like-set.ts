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
  messageId: number;
  enable?: boolean;
  force?: boolean;
}

class YammerMessageLikeSetCommand extends YammerCommand {
  public get name(): string {
    return commands.MESSAGE_LIKE_SET;
  }

  public get description(): string {
    return 'Likes or unlikes a Yammer message';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        enable: args.options.enable,
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--messageId <messageId>'
      },
      {
        option: '--enable [enable]',
        autocomplete: ['true', 'false']
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('enable');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.messageId && typeof args.options.messageId !== 'number') {
          return `${args.options.messageId} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const executeLikeAction: () => Promise<void> = async (): Promise<void> => {
      const endpoint = `${this.resource}/v1/messages/liked_by/current.json`;
      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          message_id: args.options.messageId
        }
      };

      try {
        if (args.options.enable !== false) {
          await request.post(requestOptions);
        }
        else {
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.enable === false) {
      if (args.options.force) {
        await executeLikeAction();
      }
      else {
        const messagePrompt = `Are you sure you want to unlike message ${args.options.messageId}?`;

        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: messagePrompt
        });

        if (result.continue) {
          await executeLikeAction();
        }
      }
    }
    else {
      await executeLikeAction();
    }
  }
}

export default new YammerMessageLikeSetCommand();