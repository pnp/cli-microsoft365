import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../GraphItemsListCommand';
import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import { Channel } from './Channel';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}
interface Options extends GlobalOptions {
  teamId: string;
  channelName: string;
  newChannelName: string;
  description: string
}
class GraphTeamsChannelSetCommand extends GraphItemsListCommand<Channel>{
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_SET}`;
  }
  public get description(): string {
    return 'Updates properties of a specified channel in the given Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.teamId;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let channelId: string = '';
    const endpoint: string = `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels`;
    this.getAllItems(endpoint, cmd, true)
      .then((): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Channels in current Microsoft Teams team')
          cmd.log(this.items);
          cmd.log('');
        }
        const channelItem: Channel | undefined = this.items.find(c => c.displayName && c.id ? (c.displayName.toLowerCase() === args.options.channelName.toLowerCase()) : false);
        if (channelItem) {
          if (this.debug) {
            cmd.log('The specified channel to be updated')
            cmd.log(channelItem);
            cmd.log('');
          }
          channelId = channelItem.id;
        }
        else {
          throw new Error(`The specified channel does not exist in the Microsoft Teams team`);
        }
        const endpoint: string = `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${channelId}`;
        const requestOptions: any = {
          url: endpoint,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'accept': 'application/json;odata.metadata=none'
          }),
          json: true,
          body: { "description": args.options.description, "displayName": args.options.newChannelName }
        };
        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }
        return request.patch(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        } cb();
      }, (err: any) => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team for which to update channel'
      },
      {
        option: '--channelName <channelName>',
        description: 'The name of the channel that needs to be updated'
      },
      {
        option: '--newChannelName <newChannelName>',
        description: 'The new name of the channel'
      },
      {
        option: '--description <description>',
        description: 'The description of the channel'
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
      if (!args.options.channelName) {
        return 'Required parameter channelName missing';
      }
      if (args.options.channelName.toLowerCase() === "general") {
        return 'General channel cannot be patched';
      }
      if (!args.options.newChannelName) {
        return 'Required parameter newChannelName missing';
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

  Updates properties of a specified channel in the given Microsoft Teams team,
    you have to first log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Update properties of a specified channel in the given Microsoft Teams team with description 
      ${chalk.grey(config.delimiter)} ${this.name} --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects --description "Channel for new projects"

    Update properties of a specified channel in the given Microsoft Teams team without description 
      ${chalk.grey(config.delimiter)} ${this.name} --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects
`);
  }
}
module.exports = new GraphTeamsChannelSetCommand();