import { Post } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { aadGroup } from '../../../../utils/aadGroup';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  threadId: string;
}

class AadO365GroupConversationPostListCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_CONVERSATION_POST_LIST;
  }

  public get description(): string {
    return 'Lists conversation posts of a Microsoft 365 group';
  }
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    telemetryProps.groupDisplayName = typeof args.options.groupDisplayName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['receivedDateTime', 'id'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getGroupId(args)
      .then((retrievedgroupId: string): Promise<Post[]> => {
        return odata.getAllItems<Post>(`${this.resource}/v1.0/groups/${retrievedgroupId}/threads/${args.options.threadId}/posts`);
      })
      .then((posts): void => {
        logger.log(posts);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.groupId) {
      return Promise.resolve(encodeURIComponent(args.options.groupId));
    }
    
    return aadGroup
      .getGroupByDisplayName(args.options.groupDisplayName!)
      .then(group => group.id!);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '-d, --groupDisplayName [groupDisplayName]'
      },
      {
        option: '-t, --threadId <threadId>'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.groupId &&
      !args.options.groupDisplayName) {
      return 'Specify either groupId or groupDisplayName';
    }
    if (args.options.groupId && args.options.groupDisplayName) {
      return 'Specify either groupId or groupDisplayName, but not both';
    }
    if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupConversationPostListCommand();