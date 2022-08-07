import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: number;
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
        userId: args.options.id !== undefined,
        email: args.options.email !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '--email [email]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id !== undefined && args.options.email !== undefined) {
          return `You are only allowed to search by ID or e-mail but not both`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endPoint = `${this.resource}/v1/users/current.json`;

    if (args.options.id) {
      endPoint = `${this.resource}/v1/users/${encodeURIComponent(args.options.id)}.json`;
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

    try {
      const res: any = await request.get(requestOptions);
      
      logger.log(res);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new YammerUserGetCommand();