import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  messageId: number;
  enable?: boolean;
  confirm?: boolean;
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
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
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
    if (args.options.enable === false) {
      if (args.options.confirm) {
        await this.executeLikeAction(args.options);
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
          await this.executeLikeAction(args.options);
        }
      }
    }
    else {
      await this.executeLikeAction(args.options);
    }
  }

  private async executeLikeAction(options: Options): Promise<void> {
    const endpoint = `${this.resource}/v1/messages/liked_by/current.json`;
    const requestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        message_id: options.messageId
      }
    };

    try {
      if (options.enable !== false) {
        await request.post(requestOptions);
      }
      else {
        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new YammerMessageLikeSetCommand();