import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { Team } from './Team';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { GroupUserCollection } from './groupUserCollection';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
  role?: string;
}

class TeamsListCommand extends GraphItemsListCommand<Team> {
  public get name(): string {
    return `${commands.TEAMS_USERS_LIST}`;
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
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/groups/${args.options.groupId}/owners?$select=id,displayName,userPrincipalName,userType`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((groupUsers: GroupUserCollection): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(groupUsers);
          cmd.log('');
        }

        if (args.options.output === 'json') {
          cmd.log(groupUsers);
        }
        else {
          cmd.log(groupUsers.value.map(u => {
            return {
              DisplayName: u.displayName,
              UserPrincipalName: u.userPrincipalName,
              UserType: u.userType
            };
          }));
        }

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
      if (!args.options.groupId) {
        return 'Required parameter groupId missing';
      }

      if (!Utils.isValidGuid(args.options.groupId as string)) {
        return `${args.options.groupId} is not a valid GUID`;
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

module.exports = new TeamsListCommand();