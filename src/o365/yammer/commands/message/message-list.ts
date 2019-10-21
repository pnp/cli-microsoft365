import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import YammerCommand from "../../../base/YammerCommand";
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  olderThanId?: number;
  threaded?: string;
  limit?: number;
}

class YammerMessageListCommand extends YammerCommand {
  protected items: any[];

  /* istanbul ignore next */
  constructor() {
    super();
    this.items = [];
  }

  public get name(): string {
    return `${commands.YAMMER_MESSAGE_LIST}`;
  }

  public get description(): string {
    return 'Returns all accessible messages from the user’s Yammer network.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.olderThanId = args.options.olderThanId !== undefined;
    telemetryProps.threaded  = args.options.threaded !== undefined;
    telemetryProps.limit = args.options.limit !== undefined;
    return telemetryProps;
  }

  private getAllItems(cmd: CommandInstance, args: CommandArgs, firstRun: boolean, messageId: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (firstRun) 
        this.items = [];
      let endPoint = `${this.resource}/v1/messages.json`
      
      if (messageId !== -1) 
        endPoint += `?older_than=${messageId}`
      else if (args.options.olderThanId) 
        endPoint += `?older_than=${args.options.olderThanId}`
      
      if (args.options.threaded) {
        if (endPoint.indexOf("?") > -1) 
          endPoint += "&";
        else 
          endPoint += "?"
        endPoint += `threaded=${args.options.threaded}`
      }

      const requestOptions: any = {
        url: endPoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        json: true,
        body: {
        }
      };

      request
      .get(requestOptions)
      .then((res: any): void => {
        let metadata = res.meta;
        let messageOutput = res.messages;

        cmd.log(metadata);
     
        if (args.options.output === 'json') {
          this.items = this.items.concat(messageOutput);
        }
        else {
          this.items = this.items.concat((messageOutput as any[]).map(n => {
            const item: any = {
              id: n.id, 
              sender_id: n.sender_id, 
              replied_to_id: n.replied_to_id, 
              thread_id: n.thread_id, 
              group_id: n.group_id,
              created_at: n.created_at              
            };
            return item;
          }));
        }
        
        if (args.options.limit && this.items.length > args.options.limit) {
          this.items = this.items.slice(0,args.options.limit);
          resolve();
        }
        else {
          if (metadata.older_available === true) {
            this.getAllItems(cmd, args, false, this.items[this.items.length - 1].id)
                .then((): void => {
                  resolve();
                }, (err: any): void => {
                  reject(err);
                });
          }
          else {
            resolve();
          }
        }
      }, (err: any): void => {
        reject(err);
      });
    });
  };

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
    .getAllItems(cmd, args, true, -1)
    .then((): void => {
        cmd.log(this.items);
        cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-o, --olderThanId [olderThanId]',
        description: 'Returns messages older than the message ID specified as a numeric string'
      },
      {
        option: '--threaded [threaded]',
        description: 'threaded=true will only return the thread starter (first message) for each thread. This parameter is intended for apps which need to display message threads collapsed. threaded=extended will return the thread starter messages and the two most recent messages all ordered by activity, as they are viewed in the default view on the Yammer web interface.',
        autocomplete: ['true', 'extended']
      },
      {
        option: '--limit [limit]',
        description: 'Limits the messages returned'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      ` Examples:
  
    Returns all Yammer network messages
      ${this.name}
    
    Returns all Yammer network messages older than the message ID 5611239081
      ${this.name} --olderThanId 5611239081

    Returns all Yammer network thread starter (first message) for each thread
      ${this.name} --threaded

    Returns the first 10 Yammer network messages
      ${this.name} --limit 10
    `);
  }
}

module.exports = new YammerMessageListCommand();