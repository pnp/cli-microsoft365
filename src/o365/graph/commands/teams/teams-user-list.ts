import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import Utils from '../../../../Utils';
import { GroupUser } from './GroupUser';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  role?: string;
}


class TeamsUserListCommand extends GraphItemsListCommand<GroupUser> {
  public get name(): string {
    return `${commands.TEAMS_USER_LIST}`;
  }

  public get description(): string {
    return 'Lists users of the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.getOwners(cmd, args.options.teamId)
      .then((): Promise<any> => {
        cmd.log(args.options.role);

        if (args.options.role != "Owner") {

          return this.getMembersAndGuests(cmd, args.options.teamId)
            .then((): Promise<any> => {

              // Filter out duplicate added values for owners (as they are returned as members as well)
              this.items = this.items.filter((groupUser, index, self) =>
                index === self.findIndex((t) => (
                  t.id === groupUser.id && t.displayName === groupUser.displayName
                ))
              )

              return Promise.resolve();
            });
        } else {
          return Promise.resolve();
        }
      })

      .then((): void => {
        if (args.options.role) {
          this.items = this.items.filter(i => i.userType === args.options.role)
        }

        if (this.debug) {
          cmd.log('Response:');
          cmd.log(this.items);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getOwners(cmd: CommandInstance, teamId: string): Promise<void> {
    let endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, true).then((): void => {

      // Currently there is a bug in the Graph that returns Owners as userType 'member'
      // We therefor update all returned user as owner  
      for (var i in this.items) {
        this.items[i].userType = "Owner";
      }
    });
  }

  private getMembersAndGuests(cmd: CommandInstance, teamId: string): Promise<any> {
    let endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/members?$select=id,displayName,userPrincipalName,userType`;
    return this.getAllItems(endpoint, cmd, false);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The GroupId of the team'
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
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
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

    To list  users in  Microsoft Teams, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    You can only see the users of the Microsoft Teams you are a member of.

  Examples:
  
    List all users and their role in the selected team 
      ${chalk.grey(config.delimiter)} ${this.name} --i '00000000-0000-0000-0000-000000000000'

    List all owners and their role in the selected team 
      ${chalk.grey(config.delimiter)} ${this.name} --i '00000000-0000-0000-0000-000000000000' -r Owner 

    List all guests and their role in the selected team 
      ${chalk.grey(config.delimiter)} ${this.name} --i '00000000-0000-0000-0000-000000000000' -r Guest
`);
  }
}

module.exports = new TeamsUserListCommand();