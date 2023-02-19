import { Room } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  roomlistEmail?: string;
}

class OutlookRoomListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROOM_LIST;
  }

  public get description(): string {
    return 'Get a collection of all available rooms';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'phone', 'emailAddress'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        roomlistEmail: typeof args.options.roomlistEmail !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--roomlistEmail [roomlistEmail]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let endpoint: string = `${this.resource}/v1.0/places/microsoft.graph.room`;

    if (args.options.roomlistEmail) {
      endpoint = `${this.resource}/v1.0/places/${args.options.roomlistEmail}/microsoft.graph.roomlist/rooms`;
    }

    try {
      const rooms = await odata.getAllItems<Room>(endpoint);
      await logger.log(rooms);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookRoomListCommand();
