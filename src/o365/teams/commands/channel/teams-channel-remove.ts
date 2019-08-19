import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelId: string;
  teamId: string;
  confirm?: boolean;
}

class TeamsChannelRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAMS_CHANNEL_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified channel in the Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.channelId = args.options.channelId;
    telemetryProps.teamsId = args.options.teamId;
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (args.options.confirm) {
      this.removeChannel();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the channel ${args.options.channelId} from team ${args.options.teamId}?`,
      }, (result: { continue: boolean }): void => {
        cb();
      });
    }
  }

  private removeChannel(): void {

  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --channelId <channelId>',
        description: 'The ID of the channel to remove'
      },
      {
        option: '-i, --teamId [teamId]',
        description: 'The ID of the team to which the channel to remove belongs'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.channelId) {
        return 'Required parameter channelId missing';
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId)) {
        return `${args.options.channelId} is not a valid Teams Channel Id`;
      }

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
      
      When deleted, Microsoft Teams channels are moved to a recycle bin and can be restored within 30 days. After that time, they are permanently deleted.
      
  Examples:
    
    Removes the specified Teams channel
      ${this.name} --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656

    Removes the specified Teams channel without confirmation
      ${this.name} --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656 --confirm
`);
  }
}

module.exports = new TeamsChannelRemoveCommand();