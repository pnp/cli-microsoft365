import { Conversation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
}

class AadO365GroupConversationListCommand extends GraphItemsListCommand<Conversation> {
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
    this
      .getAllItems(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`, logger, true)
      .then((): void => {
        logger.log(this.items);
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
    if (!Utils.isValidGuid(args.options.groupId as string)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupConversationListCommand();