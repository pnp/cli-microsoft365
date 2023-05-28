import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
  skipRecycleBin: boolean;
}

class AadO365GroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_REMOVE;
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
        confirm: (!(!args.options.confirm)).toString(),
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
        option: '--confirm'
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
        logger.logToStderr(`Removing Microsoft 365 Group: ${args.options.id}...`);
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

    if (args.options.confirm) {
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

module.exports = new AadO365GroupRemoveCommand();