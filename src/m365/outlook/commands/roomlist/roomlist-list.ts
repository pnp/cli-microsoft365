import { RoomList } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const roomLists = await odata.getAllItems<RoomList>(`${this.resource}/v1.0/places/microsoft.graph.roomlist`);
      logger.log(roomLists);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new OutlookRoomListListCommand();