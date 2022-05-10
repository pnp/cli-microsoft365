import { RoomList } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class OutlookRoomListListCommand extends GraphCommand {
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
    odata
      .getAllItems<RoomList>(`${this.resource}/v1.0/places/microsoft.graph.roomlist`)
      .then((roomLists): void => {
        logger.log(roomLists);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new OutlookRoomListListCommand();