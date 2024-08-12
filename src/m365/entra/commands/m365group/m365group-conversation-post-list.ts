import { Post } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupName?: string;
  threadId: string;
}

class EntraM365GroupConversationPostListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_POST_LIST;
  }

  public get description(): string {
    return 'Lists conversation posts of a Microsoft 365 group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_CONVERSATION_POST_LIST];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '-d, --groupName [groupName]'
      },
      {
        option: '-t, --threadId <threadId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupId', 'groupName'] });
  }

  public defaultProperties(): string[] | undefined {
    return ['receivedDateTime', 'id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_CONVERSATION_POST_LIST, commands.M365GROUP_CONVERSATION_POST_LIST);

    try {
      const retrievedgroupId = await this.getGroupId(args);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(retrievedgroupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${retrievedgroupId}' is not a Microsoft 365 group.`);
      }

      const posts = await odata.getAllItems<Post>(`${this.resource}/v1.0/groups/${retrievedgroupId}/threads/${args.options.threadId}/posts`);
      await logger.log(posts);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.groupId) {
      return formatting.encodeQueryParameter(args.options.groupId);
    }

    const group = await entraGroup.getGroupByDisplayName(args.options.groupName!);
    return group.id!;
  }
}

export default new EntraM365GroupConversationPostListCommand();