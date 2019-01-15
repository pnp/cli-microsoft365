import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  userName: string;
  confirm?: boolean;
}

class GraphTeamsUserRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_USER_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified user from the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let userId = '';

    const removeUser: () => void = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: { value: string; }): request.RequestPromise => {
          if (this.debug) {
            cmd.log('Response:')
            cmd.log(res);
            cmd.log('');
          }

          userId = res.value;

          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/groups/${args.options.teamId}/owners?$select=id,displayName,userPrincipalName,userType`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: any): Promise<void> | request.RequestPromise => {
          if (this.debug) {
            cmd.log('Response:')
            cmd.log(res);
            cmd.log('');
          }

          const userIsOwner: boolean = (res.value.filter((i: any) => i.userPrincipalName === args.options.userName).length > 0);

          const endpoint: string = `${auth.service.resource}/v1.0/groups/${args.options.teamId}/${userIsOwner ? 'owners' : 'members'}/${userId}/$ref`;

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
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeUser();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove ${args.options.userName} from the team ${args.options.teamId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeUser();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team from which to remove the user'
      },
      {
        option: '-n, --userName <userName>',
        description: 'User\'s UPN (user principal name), eg. johndoe@example.com'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing user from the specified team'
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

      if (!args.options.userName) {
        return 'Required parameter userName missing';
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

    To remove a user from the specified Microsoft Teams team, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    You can remove users from a Microsoft Teams team if you are owner of that
    team.

  Examples:
  
    Removes user from the specified team 
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'

    Removes user from the specified team without confirmation
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --confirm 
  `);
  }
}

module.exports = new GraphTeamsUserRemoveCommand();