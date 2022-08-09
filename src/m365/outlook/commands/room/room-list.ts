import { Room } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/v1.0/places/microsoft.graph.room`;

    if (args.options.roomlistEmail) {
      endpoint = `${this.resource}/v1.0/places/${args.options.roomlistEmail}/microsoft.graph.roomlist/rooms`;
    }

    odata
      .getAllItems<Room>(endpoint)
      .then((rooms): void => {
        logger.log(rooms);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new OutlookRoomListCommand();
