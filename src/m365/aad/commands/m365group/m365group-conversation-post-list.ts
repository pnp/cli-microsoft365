import { Post } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  threadId: string;
}

class AadM365GroupConversationPostListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_POST_LIST;
  }

  public get description(): string {
    return 'Lists conversation posts of a Microsoft 365 group';
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
        groupDisplayName: typeof args.options.groupDisplayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '-d, --groupDisplayName [groupDisplayName]'
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
    this.optionSets.push({ options: ['groupId', 'groupDisplayName'] });
  }

  public defaultProperties(): string[] | undefined {
    return ['receivedDateTime', 'id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const retrievedgroupId = await this.getGroupId(args);
      const isUnifiedGroup = await aadGroup.isUnifiedGroup(retrievedgroupId);
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

    const group = await aadGroup.getGroupByDisplayName(args.options.groupDisplayName!);
    return group.id!;
  }
}

export default new AadM365GroupConversationPostListCommand();