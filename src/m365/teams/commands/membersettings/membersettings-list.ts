import { Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class TeamsMemberSettingsListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEMBERSETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists member settings for a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId <teamId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=memberSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: Team = await request.get<Team>(requestOptions);
      logger.log(res.memberSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMemberSettingsListCommand();