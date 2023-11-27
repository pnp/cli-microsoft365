import { Group } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  force?: boolean;
}

class TeamsTeamRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
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
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const group = await aadGroup.getGroupByDisplayName(args.options.name!);
    if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
      throw `The specified team does not exist in the Microsoft Teams`;
    }

    return group.id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTeam = async (): Promise<void> => {
      try {
        const teamId: string = await this.getTeamId(args);
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(teamId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeTeam();
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the team ${args.options.id ? args.options.id : args.options.name}?` });

      if (result) {
        await removeTeam();
      }
    }
  }
}

export default new TeamsTeamRemoveCommand();