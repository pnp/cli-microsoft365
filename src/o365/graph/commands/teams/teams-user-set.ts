import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import Utils from '../../../../Utils';
import { GroupUser } from '../o365group/GroupUser';
import * as request from 'request-promise-native';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role: string;
  teamId: string;
  userName: string;
}

class GraphTeamsUserSetCommand extends GraphItemsListCommand<GroupUser> {
  public get name(): string {
    return `${commands.TEAMS_USER_SET}`;
  }

  public get description(): string {
    return 'Promote or demote the specified member or owner for the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.role;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {

    this
      .getOwners(cmd, args.options.teamId)
      .then((): Promise<void> => {
        return this.getMembersAndGuests(cmd, args.options.teamId);
      })
      .then((): request.RequestPromise | void => {

        // Filter out duplicate added values for owners (as they are returned as members as well)
        this.items = this.items.filter((groupUser, index, self) =>
          index === self.findIndex((t) => (
            t.id === groupUser.id && t.displayName === groupUser.displayName
          ))
        );

        if (this.debug) {
          cmd.log('Team owners and members:')
          cmd.log(this.items);
          cmd.log('');
        }

        if (this.items.filter(i => i.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase()).length <= 0) {
          throw new Error("The specified user is no owner or member in the specified team, please use the teams user add to add new users.");
        }

        if (args.options.role === "Owner") {
          const foundMember = this.items.find(e => e.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase() && e.userType === 'Member');

          if (foundMember !== undefined) {
            const endpoint: string = `${auth.service.resource}/v1.0/groups/${args.options.teamId}/owners/$ref`;

            const requestOptions: any = {
              url: endpoint,
              headers: Utils.getRequestHeaders({
                authorization: `Bearer ${auth.service.accessToken}`,
                'accept': 'application/json;odata.metadata=none'
              }),
              json: true,
              body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + foundMember.id }
            };

            if (this.debug) {
              cmd.log('Executing web request...');
              cmd.log(requestOptions);
              cmd.log('');
            }

            return request.post(requestOptions);
          }
          else {
            throw new Error("The specified user is already an owner in the specified team, and thus cannot be promoted.");
          }
        }
        else {
          const foundOwner = this.items.find(e => e.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase() && e.userType === 'Owner');

          if (foundOwner !== undefined) {
            const endpoint: string = `${auth.service.resource}/v1.0/groups/${args.options.teamId}/owners/${foundOwner.id}/$ref`;

            const requestOptions: any = {
              url: endpoint,
              headers: Utils.getRequestHeaders({
                authorization: `Bearer ${auth.service.accessToken}`,
                'accept': 'application/json;odata.metadata=none'
              }),
            };

            if (this.debug) {
              cmd.log('Executing web request...');
              cmd.log(requestOptions);
              cmd.log('');
            }

            return request.delete(requestOptions);
          }
          else {
            throw new Error("The specified user is already an member in the specified team, and thus cannot be demoted.");
          }
        }
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getOwners(cmd: CommandInstance, teamId: string): Promise<void> {
    const endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, true).then((): void => {
      // Currently there is a bug in the Microsoft Graph that returns Owners as
      // userType 'member'. We therefore update all returned user as owner  
      for (var i in this.items) {
        this.items[i].userType = 'Owner';
      }
    });
  }

  private getMembersAndGuests(cmd: CommandInstance, teamId: string): Promise<void> {
    const endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/members?$select=id,displayName,userPrincipalName,userType`;
    return this.getAllItems(endpoint, cmd, false);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team for which to list users'
      },
      {
        option: '-n, --userName <userName>',
        description: 'User\'s UPN (user principal name, eg. johndoe@example.com)'
      },
      {
        option: '-r, --role [type]',
        description: 'Filter the results to only users with the given role: Owner|Member',
        autocomplete: ['Owner', 'Member']
      },

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

      if (!args.options.userName) {
        return 'Required parameter userName missing';
      }

      if (!args.options.role) {
        return 'Required parameter role missing';
      } else {
        if (['Owner', 'Member'].indexOf(args.options.role) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
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

    To promote or demote members and owners in the specified Microsoft Teams team, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
    Promote the specified member to Owner  
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner

    Demote the specified member to Member  
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Member
`);
  }
}

module.exports = new GraphTeamsUserSetCommand();