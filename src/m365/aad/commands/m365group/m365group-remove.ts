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
  skipRecycleBin: boolean;
}

class AadM365GroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Microsoft 365 Group';
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
        force: (!(!args.options.force)).toString(),
        skipRecycleBin: args.options.skipRecycleBin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
      },
      {
        option: '--skipRecycleBin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing Microsoft 365 Group: ${args.options.id}...`);
      }

      try {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${args.options.id}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);

        if (args.options.skipRecycleBin) {
          const requestOptions2: CliRequestOptions = {
            url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}`,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            }
          };
          await request.delete(requestOptions2);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id}?`
      });

      if (result.continue) {
        await removeGroup();
      }
    }
  }
}

export default new AadM365GroupRemoveCommand();