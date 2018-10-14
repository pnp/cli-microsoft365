import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from "../../GraphCommand";
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Channel } from './Channel';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId : string;
  name: string;
  description?: string;
}

class TeamsChannelAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_ADD}`;
  }

  public get description(): string {
    return 'Add a channel to a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.groupId = args.options.groupId;
    telemetryProps.displayName = args.options.name;
    telemetryProps.description = args.options.description;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {

    let endpoint: string = `${auth.service.resource}/beta/teams/${args.options.groupId}/channels`;
   
    auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: endpoint,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none',
              'content-type': 'application/json;odata=nometadata'
            }),
            body: {
              displayName: args.options.name,
              description: (args.options.description) ? args.options.description : null
            },
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: any): void => {
          var channel : Channel = {
            id: res.id,
            displayName: res.displayName,
            description: res.description
          };
          cmd.log(channel);
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
        description: 'The group id to add the channel.'
      },
      {
        option: '-n, --name <name>',
        description: 'The name of the channel.'
      },
      {
        option: '-d, --description [description]',
        description: 'The description of the channel.'
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
        return `GroupId ${args.options.groupId} is not a valid GUID`;
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
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

        To add a channel top Microsoft Teams team, you have to first log in to
        the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
        eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

        You can only add a channel to the Microsoft Teams team
        you are a member of.

      Examples:
      
        Add channel to the Microsoft Teams team in the tenant
          ${chalk.grey(config.delimiter)} ${this.name} -i 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 -n office365cli -d development
`   );
  }
}

module.exports = new TeamsChannelAddCommand();