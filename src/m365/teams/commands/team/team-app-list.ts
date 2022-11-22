import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const group: any = await aadGroup.getGroupByDisplayName(args.options.name!);
    if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
      throw `The specified team does not exist in the Microsoft Teams`;
    }

    return group.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Retrieving installed apps for team ${args.options.id || args.options.name}`);
      }

      const teamId: string = await this.getTeamId(args);
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/installedApps?$expand=teamsApp,teamsAppDefinition`);

      if (!args.options.output || args.options.output === 'json') {
        logger.log(res);
      }
      else {
        //converted to text friendly output
        logger.log(res.map(i => {
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
}

module.exports = new TeamsTeamAppListCommand();