import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { RoomList } from '@microsoft/microsoft-graph-types';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';

interface CommandArgs {
  options: GlobalOptions;
}

class OutlookRoomlistListCommand extends GraphItemsListCommand<RoomList> {
  public get name(): string {
    return commands.ROOMLIST_LIST;
  }

  public get description(): string {
    return 'Get a collection of available roomlists';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'phone', 'emailAddress'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.getAllItems(`${this.resource}/v1.0/places/microsoft.graph.roomlist`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new OutlookRoomlistListCommand();