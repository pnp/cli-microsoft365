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
  userId?: number;
  email?: string;
}

class YammerGroupUserAddCommand extends YammerCommand {
  public get name(): string {
    return commands.GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Adds a user to a Yammer Group';
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
        userId: typeof args.options.userId !== 'undefined',
        email: typeof args.options.email !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--email [email]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        if (args.options.userId && typeof args.options.userId !== 'number') {
          return `${args.options.userId} is not a number`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/group_memberships.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        group_id: args.options.id,
        user_id: args.options.userId,
        email: args.options.email
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new YammerGroupUserAddCommand();