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

  // assumption we can only have 100 owners per team, but 2500 members
  // do call to /members & /owners if no paramter is added
  // do call to /members for guest and members (filter on usertype)

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let groupOwners: GroupUser[] = [];
    let groupMembers: GroupUser[] = [];
    //let users: GroupUser[] = [];

    this.getOwners(cmd, args.options.teamId)
      .then((): Promise<any> => {
        console.log("Retrieved owners");

        // currently there is a bug in the Graph that returns Owners as userType 'member'
        // We therefor update all returned user as owner  
        for (var i in this.items) {
          this.items[i].userType = "Owner";
        }

        groupOwners = this.items;

        if (args.options.role || args.options.role !== "Owner") {
          console.log("retrieving Members");
          return this.getMembers(cmd, args.options.teamId)
            .then((): Promise<any> => {

              groupMembers = this.items;

              return Promise.resolve();
            });
        } else {
          return Promise.resolve();
        }
      })

      .then((): void => {

        // construct array with both members and owners / filter out duplicate 

        if (this.debug) {
          cmd.log('Response:')
          cmd.log(groupOwners)
          cmd.log(groupMembers)
          cmd.log('');
        }

        //cmd.log(groupOwners);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getOwners(cmd: CommandInstance, teamId: string): Promise<any> {
    cmd.log('Retrieving Owners')
    let endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/owners?$select=id,displayName,userPrincipalName,userType`;

    // Bug in the graph, we should upgrade to owners
    return this.getAllItems(endpoint, cmd, true);
  }

  private getMembers(cmd: CommandInstance, teamId: string): Promise<any> {
    cmd.log('Retrieving Members')

    let endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/members?$select=id,displayName,userPrincipalName,userType`;
    return this.getAllItems(endpoint, cmd, true);
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