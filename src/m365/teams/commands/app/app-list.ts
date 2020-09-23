import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { TeamsApp } from '../../TeamsApp';
import { TeamsAppInstallation } from '../../TeamsAppInstallation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  all?: boolean;
  teamId?: string;
}

class TeamsAppListCommand extends GraphItemsListCommand<TeamsApp> {
  public get name(): string {
    return `${commands.TEAMS_APP_LIST}`;
  }

  public get description(): string {
    return 'Lists apps from the Microsoft Teams app catalog or apps installed in the specified team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.all = args.options.all || false;
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let endpoint: string = '';
    if (args.options.teamId) {
      endpoint = `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/installedApps?$expand=teamsApp`;

      if (!args.options.all) {
        endpoint += `&$filter=teamsApp/distributionMethod eq 'organization'`;
      }
    }
    else {
      endpoint = `${this.resource}/v1.0/appCatalogs/teamsApps`;

      if (!args.options.all) {
        endpoint += `?$filter=distributionMethod eq 'organization'`;
      }
    }

    this
      .getAllItems(endpoint, logger, true)
      .then((): void => {
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          if (args.options.teamId) {
            logger.log((this.items as unknown as TeamsAppInstallation[]).map(i => {
              return {
                id: i.id,
                displayName: i.teamsApp.displayName,
                distributionMethod: i.teamsApp.distributionMethod
              };
            }));
          }
          else {
            logger.log(this.items.map(i => {
              return {
                id: i.id,
                displayName: i.displayName,
                distributionMethod: i.distributionMethod
              };
            }));
          }
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-a, --all',
        description: 'Specify, to get apps from your organization only'
      },
      {
        option: '-i, --teamId [teamId]',
        description: 'The ID of the team for which to list installed apps'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.teamId && !Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsAppListCommand();