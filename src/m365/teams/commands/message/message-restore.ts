import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand.js';
import { teams } from '../../../../utils/teams.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  id: string;
}

class TeamsMessageRestoreCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.MESSAGE_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted message from a channel in a Microsoft Teams team';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '--channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `'${args.options.teamId}' is not a valid GUID for 'teamId'.`;
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
          return `'${args.options.channelId}' is not a valid ID for 'channelId'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] });
  }

  #initTypes(): void {
    this.types.string.push('teamId', 'teamName', 'channelId', 'channelName', 'id');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Restoring deleted message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in the Microsoft Teams team '${args.options.teamId || args.options.teamName}'.`);
      }

      const teamId: string = await this.getTeamId(args.options, logger);
      const channelId: string = await this.getChannelId(args.options, teamId, logger);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${teamId}/channels/${channelId}/messages/${args.options.id}/undoSoftDelete`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(options: Options, logger: Logger): Promise<string> {
    if (options.teamId) {
      return options.teamId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Getting the Team ID.`);
    }

    const groupId = await teams.getTeamIdByDisplayName(options.teamName!);

    return groupId;
  }

  private async getChannelId(options: Options, teamId: string, logger: Logger): Promise<string> {
    if (options.channelId) {
      return options.channelId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Getting the channel ID.`);
    }

    const channelId = await teams.getChannelIdByDisplayName(teamId, options.channelName!);
    return channelId;
  }
}

export default new TeamsMessageRestoreCommand();