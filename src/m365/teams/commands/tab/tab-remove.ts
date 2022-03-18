import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
    return commands.TAB_REMOVE;
  }

  public get description(): string {
    return "Removes a tab from the specified channel";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!!args.options.confirm).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeTab: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${args.options.channelId}/tabs/${encodeURIComponent(args.options.tabId)}`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };
      request.delete(requestOptions).then(_ => cb(),
        (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb)
      );
    };
    if (args.options.confirm) {
      removeTab();
    }
    else {
      Cli.prompt(
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
        option: "-i, --teamId <teamId>"
      },
      {
        option: "-c, --channelId <channelId>"
      },
      {
        option: "-t, --tabId <tabId>"
      },
      {
        option: "--confirm"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (!validation.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    if (!validation.isValidGuid(args.options.tabId as string)) {
      return `${args.options.tabId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsTabRemoveCommand();