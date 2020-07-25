import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import { Team } from '../../Team';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import request from '../../../../request';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  joined?: boolean;
}

class TeamsListCommand extends GraphItemsListCommand<Team> {
  public get name(): string {
    return `${commands.TEAMS_TEAM_LIST}`;
  }

  public get description(): string {
    return 'Lists Microsoft Teams in the current tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.joined = args.options.joined;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`;
    if (args.options.joined) {
      endpoint = `${this.resource}/beta/me/joinedTeams`;
    }
    this
      .getAllItems(endpoint, cmd, true)
      .then((): Promise<any> => {
        if (args.options.joined) {
          return Promise.resolve();
        } else {
          return Promise.all(this.items.map(g => this.getTeamFromGroup(g, cmd)));
        }
      })
      .then((res?: Team[]): void => {
        if (res) {
          this.items = res;
        }

        cmd.log(this.items);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getTeamFromGroup(group: { id: string, displayName: string, description: string }, cmd: CommandInstance): Promise<Team> {
    return new Promise<Team>((resolve: (team: Team) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/beta/teams/${group.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          resolve({
            id: group.id,
            displayName: group.displayName,
            isArchived: res.isArchived,
            description: group.description
          });
        }, (err: any): void => {
          // If the user is not member of the team he/she cannot access it
          if (err.statusCode === 403) {
            resolve({
              id: group.id,
              displayName: group.displayName,
              description: group.description,
              isArchived: undefined
            });
          }
          else {
            reject(err);
          }
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-j, --joined',
        description: 'Retrieve only joined teams'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TeamsListCommand();