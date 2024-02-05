import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  force?: boolean;
}

class EntraGroupSettingRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_REMOVE;
  }

  public get description(): string {
    return 'Removes the particular group setting';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTING_REMOVE];
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
        force: (!(!args.options.force)).toString()
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
    this.showDeprecationWarning(logger, aadCommands.GROUPSETTING_REMOVE, commands.GROUPSETTING_REMOVE);

    const removeGroupSetting = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing group setting: ${args.options.id}...`);
      }

      try {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroupSetting();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group setting ${args.options.id}?` });

      if (result) {
        await removeGroupSetting();
      }
    }
  }
}

export default new EntraGroupSettingRemoveCommand();