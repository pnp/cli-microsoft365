import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { Conversation } from '@microsoft/microsoft-graph-types';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
}

class AadO365GroupConversationListCommand extends GraphItemsListCommand<Conversation> {
  public get name(): string {
    return commands.O365GROUP_CONVERSATION_LIST;
  }

  public get description(): string {
    return 'Retrieve conversations for the specified group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    return telemetryProps;
  }
  public defaultProperties(): string[] | undefined {
    return ['topic', 'lastDeliveredDateTime', 'id'];
  }
  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.getAllItems(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`, logger, true)
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