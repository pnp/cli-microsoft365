import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: number;
  email?: string;
}

class YammerUserGetCommand extends YammerCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Retrieves the current user or searches for a user by ID or e-mail';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'full_name', 'email', 'job_title', 'state', 'url'];
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
        userId: args.options.userId !== undefined,
        email: args.options.email !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --userId [userId]'
      },
      {
        option: '--email [email]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId !== undefined && args.options.email !== undefined) {
          return `You are only allowed to search by ID or e-mail but not both`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endPoint = `${this.resource}/v1/users/current.json`;

    if (args.options.userId) {
      endPoint = `${this.resource}/v1/users/${encodeURIComponent(args.options.userId)}.json`;
    }
    else if (args.options.email) {
      endPoint = `${this.resource}/v1/users/by_email.json?email=${encodeURIComponent(args.options.email)}`;
    }

    const requestOptions: any = {
      url: endPoint,
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

module.exports = new YammerUserGetCommand();