import { Conversation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    odata
      .getAllItems<Conversation>(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`)
      .then((conversations): void => {
        logger.log(conversations);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --groupId <groupId>'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.groupId as string)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupConversationListCommand();