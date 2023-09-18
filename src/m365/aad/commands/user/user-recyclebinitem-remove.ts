import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  force?: boolean;
}

class AadUserRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes a user from the recycle bin in the current tenant';
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
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearRecycleBinItem: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Permanently deleting user with id ${args.options.id} from Azure Active Directory`);
      }

      try {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}`,
          headers: {}
        };
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearRecycleBinItem();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to permanently delete the user with id ${args.options.id}?`);

      if (result) {
        await clearRecycleBinItem();
      }
    }
  }
}

export default new AadUserRecycleBinItemRemoveCommand();
