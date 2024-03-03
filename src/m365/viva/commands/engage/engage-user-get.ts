import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: number;
  email?: string;
}

class VivaEngageUserGetCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_USER_GET;
  }

  public get description(): string {
    return 'Retrieves the current user or searches for a user by ID or e-mail';
  }

  public alias(): string[] | undefined {
    return [yammerCommands.USER_GET];
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
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    let endPoint = `${this.resource}/v1/users/current.json`;

    if (args.options.id) {
      endPoint = `${this.resource}/v1/users/${args.options.id}.json`;
    }
    else if (args.options.email) {
      endPoint = `${this.resource}/v1/users/by_email.json?email=${formatting.encodeQueryParameter(args.options.email)}`;
    }

    const requestOptions: CliRequestOptions = {
      url: endPoint,
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

export default new VivaEngageUserGetCommand();