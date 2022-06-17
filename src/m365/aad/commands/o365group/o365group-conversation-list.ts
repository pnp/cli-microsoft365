import { Conversation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
}

class AadO365GroupConversationListCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_CONVERSATION_LIST;
  }

  public get description(): string {
    return 'Lists conversations for the specified Microsoft 365 group';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    odata
      .getAllItems<Conversation>(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`)
      .then((conversations): void => {
        logger.log(conversations);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadO365GroupConversationListCommand();