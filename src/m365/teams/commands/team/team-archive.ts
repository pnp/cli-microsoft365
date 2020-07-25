import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  shouldSetSpoSiteReadOnlyForMembers: boolean;
}

class TeamsArchiveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_ARCHIVE}`;
  }

  public get description(): string {
    return 'Archives specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.shouldSetSpoSiteReadOnlyForMembers = args.options.shouldSetSpoSiteReadOnlyForMembers === true;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const siteReadOnlyForMembers: boolean = args.options.shouldSetSpoSiteReadOnlyForMembers === true;
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/archive`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      json: true,
      body: {
        shouldSetSpoSiteReadOnlyForMembers: siteReadOnlyForMembers
      }
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to archive'
      },
      {
        option: '--shouldSetSpoSiteReadOnlyForMembers',
        description: 'Sets the permissions for team members to read-only on the SharePoint Online site associated with the team'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new TeamsArchiveCommand();