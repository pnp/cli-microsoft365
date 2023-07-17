import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
}

class TeamsTeamAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_APP_LIST;
  }

  public get description(): string {
    return 'List apps installed in the specified team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
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
        teamName: typeof args.options.teamName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '-n, --teamName [teamName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving installed apps for team '${args.options.teamId || args.options.teamName}'`);
      }

      const teamId: string = await this.getTeamId(args, logger);
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`);

      if (!Cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res);
      }
      else {
        //converted to text friendly output
        await logger.log(res.map(i => {
          return {
            id: i.id,
            displayName: i.teamsApp.displayName,
            distributionMethod: i.teamsApp.distributionMethod
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const team = await teams.getTeamByName(args.options.teamName!, logger, this.verbose);
    return team.id!;
  }
}

export default new TeamsTeamAppListCommand();