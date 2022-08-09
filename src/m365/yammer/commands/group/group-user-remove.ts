import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from "../../../base/YammerCommand";
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  userId?: number;
  confirm?: boolean;
}

class YammerGroupUserRemoveCommand extends YammerCommand {
  public get name(): string {
    return commands.GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a user from a Yammer group';
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
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
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
    const executeRemoveAction: () => void = (): void => {
      const endpoint = `${this.resource}/v1/group_memberships.json`;

      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          group_id: args.options.id,
          user_id: args.options.userId
        }
      };

      request
        .delete(requestOptions)
        .then((): void => cb(),
          (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      executeRemoveAction();
    }
    else {
      let messagePrompt: string = `Are you sure you want to leave group ${args.options.id}?`;
      if (args.options.userId) {
        messagePrompt = `Are you sure you want to remove the user ${args.options.userId} from the group ${args.options.id}?`;
      }

      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: messagePrompt
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          executeRemoveAction();
        }
      });
    }
  }
}

module.exports = new YammerGroupUserRemoveCommand();