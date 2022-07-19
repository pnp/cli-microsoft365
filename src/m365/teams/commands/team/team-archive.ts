import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  teamId?: string;
  shouldSetSpoSiteReadOnlyForMembers: boolean;
}

class TeamsTeamArchiveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_ARCHIVE;
  }

  public get description(): string {
    return 'Archives specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.shouldSetSpoSiteReadOnlyForMembers = args.options.shouldSetSpoSiteReadOnlyForMembers === true;
    return telemetryProps;
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.name!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.teamId) {
      args.options.id = args.options.teamId;

      this.warn(logger, `Option 'teamId' is deprecated. Please use 'id' instead.`);
    }

    const siteReadOnlyForMembers: boolean = args.options.shouldSetSpoSiteReadOnlyForMembers === true;

    this
      .getTeamId(args)
      .then((teamId: string): Promise<void> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/archive`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            shouldSetSpoSiteReadOnlyForMembers: siteReadOnlyForMembers
          }
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--shouldSetSpoSiteReadOnlyForMembers'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id && !args.options.name && !args.options.teamId) {
      return 'Specify either id or name';
    }

    if (args.options.name && (args.options.id || args.options.teamId)) {
      return 'Specify either id or name but not both';
    }

    if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (args.options.id && !validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsTeamArchiveCommand();