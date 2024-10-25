import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  entraGroupId?: string;
  force?: boolean
}

class VivaEngageCommunityRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_REMOVE;
  }
  public get description(): string {
    return 'Removes a community';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: args.options.id !== 'undefined',
        displayName: args.options.displayName !== 'undefined',
        entraGroupId: args.options.entraGroupId !== 'undefined',
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
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `${args.options.entraGroupId} is not a valid GUID for the option 'entraGroupId'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'displayName', 'entraGroupId']
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName', 'entraGroupId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeCommunity = async (): Promise<void> => {
      try {
        let communityId = args.options.id;

        if (args.options.displayName) {
          communityId = (await vivaEngage.getCommunityByDisplayName(args.options.displayName, ['id'])).id;
        }
        else if (args.options.entraGroupId) {
          communityId = (await vivaEngage.getCommunityByEntraGroupId(args.options.entraGroupId, ['id'])).id;
        }

        if (args.options.verbose) {
          await logger.logToStderr(`Removing Viva Engage community with ID ${communityId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
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
      await removeCommunity();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove Viva Engage community '${args.options.id || args.options.displayName || args.options.entraGroupId }'?` });

      if (result) {
        await removeCommunity();
      }
    }
  }
}

export default new VivaEngageCommunityRemoveCommand();