import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { validation } from '../../../../utils/validation.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  force?: boolean
}

class AadGroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Entra ID group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUP_REMOVE];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTelemetry();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: args.options.id !== 'undefined',
        displayName: args.options.displayName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'displayName']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID for option id.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing group ${args.options.id || args.options.displayName}...`);
      }

      try {
        let groupId = args.options.id;

        if (args.options.displayName) {
          groupId = await aadGroup.getGroupIdByDisplayName(args.options.displayName);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove group '${args.options.id || args.options.displayName}'?` });

      if (result) {
        await removeGroup();
      }
    }
  }
}

export default new AadGroupRemoveCommand();