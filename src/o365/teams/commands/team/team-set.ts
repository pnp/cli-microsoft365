import config from '../../../../config';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { GroupUpdateService, Options} from '../../../aad/services/GroupUpdateService';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface TeamSetOptions extends Options{
  teamId: string;
  visibility?:string;
  imagePath?: string;
}
interface CommandArgs {
  options: TeamSetOptions;
}

class TeamsSetCommand extends GraphCommand {
  private static props: string[] = [
    'displayName',
    'description',
    'mailNickName',
    'classification',
    'visibility'
  ];

  public get name(): string {
    return `${commands.TEAMS_TEAM_SET}`;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    TeamsSetCommand.props.forEach((p: string) => {
      telemetryProps[p] = typeof (args.options as any)[p] !== 'undefined';
    });
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.MapTeamOptionProperties(args.options);
    GroupUpdateService.UpdateGroup(cmd, this.resource, args.options, this.verbose, this.debug, cb, this.handleRejectedODataJsonPromise);
  }

  private MapTeamOptionProperties(options: TeamSetOptions): void{
    options.id = options.teamId;
    options.isPrivate = options.visibility;
    options.logoPath = options.imagePath;
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team for which to update settings'
      },
      {
        option: '--displayName [displayName]',
        description: 'The display name for the Microsoft Teams team'
      },
      {
        option: '--description [description]',
        description: 'The description for the Microsoft Teams team'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of Microsoft Teams team owners to add'
      },
      {
        option: '--members [members]',
        description: 'Comma-separated list of Microsoft Teams team members to add'
      },
      {
        option: '--mailNickName [mailNickName]',
        description: 'The mail alias for the Microsoft Teams team'
      },
      {
        option: '--classification [classification]',
        description: 'The classification for the Microsoft Teams team'
      },
      {
        option: '--visibility [visibility]',
        description: 'The visibility of the Microsoft Teams team. Valid values Private|Public',
        autocomplete: ['Private', 'Public']
      },
      {
        option: '--imagePath [imagePath]',
        description: 'Path to the image file to set as the Microsoft Teams team picture'
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

      // if (args.options.visibility) {
      //   if (args.options.visibility.toLowerCase() !== 'private' && args.options.visibility.toLowerCase() !== 'public') {
      //     return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
      //   }
      // }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  
  Examples:
  
    Set Microsoft Teams team visibility as Private
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --visibility Private

    Set Microsoft Teams team classification as MBI
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --classification MBI
`);
  }
}

module.exports = new TeamsSetCommand();