import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from "../../../base/YammerCommand";
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: number;
  id?: number;
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
        userId: args.options.id !== undefined,
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
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
    const executeRemoveAction: () => Promise<void> = async (): Promise<void> => {
      const endpoint = `${this.resource}/v1/group_memberships.json`;

      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          group_id: args.options.groupId,
          user_id: args.options.id
        }
      };

      try {
        await request.delete(requestOptions);
      } 
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await executeRemoveAction();
    }
    else {
      let messagePrompt: string = `Are you sure you want to leave group ${args.options.groupId}?`;
      if (args.options.id) {
        messagePrompt = `Are you sure you want to remove the user ${args.options.id} from the group ${args.options.groupId}?`;
      }

      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: messagePrompt
      });

      if (result.continue) {
        await executeRemoveAction();
      }
    }
  }
}

module.exports = new YammerGroupUserRemoveCommand();