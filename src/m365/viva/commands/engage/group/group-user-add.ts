import { Logger } from '../../../../../cli/Logger.js';
import GlobalOptions from '../../../../../GlobalOptions.js';
import request from '../../../../../request.js';
import YammerCommand from '../../../../base/YammerCommand.js';
import commands from '../../../commands.js';
import yammerCommands from '../../../../yammer/commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: number;
  id?: number;
  email?: string;
}

class YammerGroupUserAddCommand extends YammerCommand {
  public get name(): string {
    return commands.ENGAGE_GROUP_USER_ADD;
  }

  public alias(): string[] {
    return [yammerCommands.GROUP_USER_ADD];
  }

  public get description(): string {
    return 'Adds a user to a Viva Engage Group';
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
        id: typeof args.options.id !== 'undefined',
        email: typeof args.options.email !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--groupId <groupId>'
      },
      {
        option: '--id [id]'
      },
      {
        option: '--email [email]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (typeof args.options.groupId !== 'number') {
          return `${args.options.groupId} is not a number`;
        }

        if (args.options.id && typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${this.resource}/v1/group_memberships.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        group_id: args.options.groupId,
        user_id: args.options.id,
        email: args.options.email
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new YammerGroupUserAddCommand();