import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
    if (args.options.enable === false) {
      if (args.options.force) {
        await this.executeLikeAction(args.options);
      }
      else {
        const message = `Are you sure you want to unlike message ${args.options.messageId}?`;

        const result = await Cli.promptForConfirmation({ message });

        if (result) {
          await this.executeLikeAction(args.options);
        }
      }
    }
    else {
      await this.executeLikeAction(args.options);
    }
  }

  private async executeLikeAction(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1/messages/liked_by/current.json`,
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

export default new YammerMessageLikeSetCommand();