import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
}

class VivaEngageMessageGetCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_MESSAGE_GET;
  }

  public get description(): string {
    return 'Returns a Viva Engage message';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'sender_id', 'replied_to_id', 'thread_id', 'group_id', 'created_at', 'direct_message', 'system_message', 'privacy', 'message_type', 'content_excerpt'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
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
    const requestOptions: any = {
      url: `${this.resource}/v1/messages/${args.options.id}.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageMessageGetCommand();
