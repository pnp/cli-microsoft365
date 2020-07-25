import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  tabId: string;
  confirm?: boolean;
}

class TeamsTabRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAMS_TAB_REMOVE;
  }

  public get description(): string {
    return "Removes a tab from the specified channel";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!!args.options.confirm).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const removeTab: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${args.options.channelId}/tabs/${encodeURIComponent(args.options.tabId)}`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        json: true
      };
      request.delete(requestOptions).then(
        (): void => {
          if (this.verbose) {
            cmd.log(chalk.green("DONE"));
          }
          cb();
        },
        (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb)
      );
    };
    if (args.options.confirm) {
      removeTab();
    }
    else {
      cmd.prompt(
        {
          type: "confirm",
          name: "continue",
          default: false,
          message: `Are you sure you want to remove the tab with id ${args.options.tabId} from channel ${args.options.channelId} in team ${args.options.teamId}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeTab();
          }
        }
      );
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --teamId <teamId>",
        description: "The ID of the team where the tab exists"
      },
      {
        option: "-c, --channelId <channelId>",
        description: "The ID of the channel to remove the tab from"
      },
      {
        option: "-t, --tabId <tabId>",
        description: "The ID of the tab to remove"
      },
      {
        option: "--confirm",
        description: "Don't prompt for confirmation"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
        return `${args.options.channelId} is not a valid Teams ChannelId`;
      }

      if (!Utils.isValidGuid(args.options.tabId as string)) {
        return `${args.options.tabId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new TeamsTabRemoveCommand();