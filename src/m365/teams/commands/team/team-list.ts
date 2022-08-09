import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Team } from '../../Team';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`;
    if (args.options.joined) {
      endpoint = `${this.resource}/v1.0/me/joinedTeams`;
    }
    odata
      .getAllItems<Group>(endpoint)
      .then((items): Promise<Group[] | Team[]> => {
        if (args.options.joined) {
          return Promise.resolve(items);
        }
        else {
          return Promise.all(
            items.filter((e: any) => {
              return e.resourceProvisioningOptions.indexOf('Team') > -1;
            }).map(
              g => this.getTeamFromGroup(g)
            )
          );
        }
      })
      .then((items: Group[] | Team[]): void => {
        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getTeamFromGroup(group: Group): Promise<Team> {
    return new Promise<Team>((resolve: (team: Team) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${group.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          resolve({
            id: group.id as string,
            displayName: group.displayName as string,
            isArchived: res.isArchived,
            description: group.description as string
          });
        }, (err: any): void => {
          // If the user is not member of the team he/she cannot access it
          if (err.statusCode === 403) {
            resolve({
              id: group.id as string,
              displayName: group.displayName as string,
              description: group.description as string,
              isArchived: undefined
            });
          }
          else {
            reject(err);
          }
        });
    });
  }
}

module.exports = new TeamsTeamListCommand();