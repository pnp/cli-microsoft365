import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
}

class YammerMessageGetCommand extends YammerCommand {
  public get name(): string {
    return commands.MESSAGE_GET;
  }

  public get description(): string {
    return 'Returns a Yammer message';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/messages/${args.options.id}.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new YammerMessageGetCommand();
