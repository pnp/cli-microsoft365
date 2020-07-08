import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    Using this command, global admins and Microsoft Teams service admins can
    access teams that they are not a member of.

    When a team is archived, users can no longer send or like messages on any
    channel in the team, edit the team\'s name, description, or other settings,
    or in general make most changes to the team. Membership changes to the team
    continue to be allowed.

  Examples:
    
    Archive the specified Microsoft Teams team
      ${commands.TEAMS_TEAM_ARCHIVE} --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
    
    Archive the specified Microsoft Teams team and set permissions for team
    members to read-only on the SharePoint Online site associated with the team
      ${commands.TEAMS_TEAM_ARCHIVE} --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --shouldSetSpoSiteReadOnlyForMembers
    `);
  }
}

module.exports = new TeamsArchiveCommand();