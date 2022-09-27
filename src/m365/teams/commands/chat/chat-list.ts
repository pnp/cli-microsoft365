import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
}

class TeamsChatListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_LIST;
  }

  public get description(): string {
    return 'Lists all chat conversations';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'topic', 'chatType'];
  }

  constructor() {
    super();
  
    
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: args.options.type
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: ['oneOnOne', 'group', 'meeting']
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const supportedTypes = ['oneOnOne', 'group', 'meeting'];
        if (args.options.type !== undefined && supportedTypes.indexOf(args.options.type) === -1) {
          return `${args.options.type} is not a valid chatType. Accepted values are ${supportedTypes.join(', ')}`;
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const filter = args.options.type !== undefined ? `?$filter=chatType eq '${args.options.type}'` : '';
    const endpoint: string = `${this.resource}/v1.0/chats${filter}`;

    try {
      const items = await odata.getAllItems(endpoint);
      logger.log(items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChatListCommand();