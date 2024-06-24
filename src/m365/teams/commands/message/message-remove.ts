import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  id: string;
  force?: boolean;
}

class TeamsMessageRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a message from a channel in a Microsoft Teams team';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined',
        force: !!args.options.force
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
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId as string)) {
          return `${args.options.channelId} is not a valid Teams ChannelId`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['teamId', 'teamName']
      },
      {
        options: ['channelId', 'channelName']
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)) {
      throw 'This command does not support application only premissions.';
    }

    const removeTeamMessage = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing message ${args.options.id} from team ${args.options.teamId || args.options.teamName} and channel ${args.options.channelId || args.options.channelName}.`);
        }

        const teamId: string = await teams.getTeamId(args.options.teamId, args.options.teamName);
        const channelId: string = await teams.getChannelId(teamId, args.options.channelId, args.options.channelName);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${args.options.id}/softdelete`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        if (err.error && err.error.error && err.error.error.code === 'NotFound') {
          this.handleError('The specified message was not found in the specified channel');
        }
        else {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await removeTeamMessage();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the message?` });

      if (result) {
        await removeTeamMessage();
      }
    }
  }
}

export default new TeamsMessageRemoveCommand();