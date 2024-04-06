import { Conversation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
}

class EntraM365GroupConversationListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_LIST;
  }

  public get description(): string {
    return 'Lists conversations for the specified Microsoft 365 group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_CONVERSATION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['topic', 'lastDeliveredDateTime', 'id'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId <groupId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_CONVERSATION_LIST, commands.M365GROUP_CONVERSATION_LIST);

    try {
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(args.options.groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.groupId}' is not a Microsoft 365 group.`);
      }

      const conversations = await odata.getAllItems<Conversation>(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`);
      await logger.log(conversations);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupConversationListCommand();