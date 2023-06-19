import { Group, Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  joined?: boolean;
}

class TeamsTeamListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Teams in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isArchived', 'description'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        joined: args.options.joined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-j, --joined'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endpoint: string = `${this.resource}/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`;
    if (args.options.joined) {
      endpoint = `${this.resource}/v1.0/me/joinedTeams`;
    }

    try {
      const items = await odata.getAllItems<Group>(endpoint);

      if (args.options.joined) {
        logger.log(items);
      }
      else {
        const teamItems = await Promise.all(
          items.filter((e: any) => {
            return e.resourceProvisioningOptions.indexOf('Team') > -1;
          }).map(
            g => this.getTeamFromGroup(g)
          )
        );
        logger.log(teamItems);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamFromGroup(group: Group): Promise<Team> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${group.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    try {
      return await request.get<Team>(requestOptions);
    }
    catch (err: any) {
      if (err.statusCode === 403) {
        return {
          id: group.id as string,
          displayName: group.displayName as string,
          description: group.description as string,
          isArchived: undefined
        };
      }
      else {
        throw err;
      }
    }
  }
}

module.exports = new TeamsTeamListCommand();