import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Tab } from '../../Tab';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
}

class TeamsTabListCommand extends GraphCommand {
  public get name(): string {
    return commands.TAB_LIST;
  }

  public get description(): string {
    return 'Lists tabs in the specified Microsoft Teams channel';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'teamsAppTabId'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${encodeURIComponent(args.options.channelId)}/tabs?$expand=teamsApp`;

    odata
      .getAllItems<Tab>(endpoint, logger)
      .then((items): void => {
        items.forEach(i => {
          (i as any).teamsAppTabId = i.teamsApp.id;
        });

        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '-c, --channelId <channelId>'
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

    return true;
  }
}

module.exports = new TeamsTabListCommand();