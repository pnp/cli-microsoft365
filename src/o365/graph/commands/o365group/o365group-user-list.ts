import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphUsersListCommand } from '../GraphUsersListCommand';
import Utils from '../../../../Utils';
import { GroupUser } from '../o365group/GroupUser';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role?: string;
  groupId: string;
}

class GraphO365GroupUserListCommand extends GraphUsersListCommand<GroupUser> {
  public get name(): string {
    return `${commands.O365GROUP_USER_LIST}`;
  }

  public get description(): string {
    return 'Lists users for the specified Office 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.role;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getOwners(cmd, args.options.groupId)
      .then((): Promise<void> => {
        if (args.options.role === 'Owner') {
          return Promise.resolve();
        }

        return this.getMembersAndGuests(cmd, args.options.groupId);
      })
      .then((): void => {
        if (args.options.role) {
          this.items = this.items.filter(i => i.userType === args.options.role)
        }

        cmd.log(this.items);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --groupId <groupId>',
        description: 'The ID of the group for which to list users'
      },
      {
        option: '-r, --role [type]',
        description: 'Filter the results to only users with the given role: Owner|Member|Guest',
        autocomplete: ['Owner', 'Member', 'Guest']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.groupId) {
        return 'Required parameter groupId missing';
      }

      if (!Utils.isValidGuid(args.options.groupId as string)) {
        return `${args.options.groupId} is not a valid GUID`;
      }

      if (args.options.role) {
        if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To list users in the specified Office 365 Group, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    List all users and their role in the specified Office 365 group 
      ${chalk.grey(config.delimiter)} ${this.name} --groupId '00000000-0000-0000-0000-000000000000'

    List all owners and their role in the specified Office 365 group 
      ${chalk.grey(config.delimiter)} ${this.name} --groupId '00000000-0000-0000-0000-000000000000' --role Owner 

    List all guests and their role in the specified Office 365 group 
      ${chalk.grey(config.delimiter)} ${this.name} --groupId '00000000-0000-0000-0000-000000000000' --role Guest
`);
  }
}

module.exports = new GraphO365GroupUserListCommand();