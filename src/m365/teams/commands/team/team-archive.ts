import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const siteReadOnlyForMembers: boolean = args.options.shouldSetSpoSiteReadOnlyForMembers === true;
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/archive`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        shouldSetSpoSiteReadOnlyForMembers: siteReadOnlyForMembers
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--shouldSetSpoSiteReadOnlyForMembers'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsArchiveCommand();