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
  threaded?: boolean;
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
    return 'Returns all accessible messages from the user\'s Yammer network';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.olderThanId = args.options.olderThanId !== undefined;
    telemetryProps.threaded = args.options.threaded;
    telemetryProps.limit = args.options.limit !== undefined;
    return telemetryProps;
  }

  private getAllItems(cmd: CommandInstance, args: CommandArgs, messageId: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      let endPoint = `${this.resource}/v1/messages.json`;

      if (messageId !== -1) {
        endPoint += `?older_than=${messageId}`;
      }
      else if (args.options.olderThanId) {
        endPoint += `?older_than=${args.options.olderThanId}`;
      }

      if (args.options.threaded) {
        if (endPoint.indexOf("?") > -1) {
          endPoint += "&";
        } else {
          endPoint += "?";
        }

        endPoint += `threaded=true`;
      }

      const requestOptions: any = {
        url: endPoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          this.items = this.items.concat(res.messages);

          if (args.options.limit && this.items.length > args.options.limit) {
            this.items = this.items.slice(0, args.options.limit);
            resolve();
          }
          else {
            if (res.meta.older_available === true) {
              this.getAllItems(cmd, args, this.items[this.items.length - 1].id)
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
    this.items = []; // this will reset the items array in interactive mode

    this
      .getAllItems(cmd, args, -1)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map((n: any) => {
            let shortBody;
            const bodyToProcess = n.body.plain;

            if (bodyToProcess) {
              let maxLength = 35;
              let addedDots = "...";
              if (bodyToProcess.length < maxLength) {
                maxLength = bodyToProcess.length;
                addedDots = "";
              }

              shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
            }

            const item: any = {
              id: n.id,
              replied_to_id: n.replied_to_id,
              thread_id: n.thread_id,
              group_id: n.group_id,
              shortBody: shortBody
            };
            return item;
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--olderThanId [olderThanId]',
        description: 'Returns messages older than the message ID specified as a numeric string'
      },
      {
        option: '--threaded',
        description: 'Will only return the thread starter (first message) for each thread. This parameter is intended for apps which need to display message threads collapsed'
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
      if (args.options.olderThanId && typeof args.options.olderThanId !== 'number') {
        return `${args.options.olderThanId} is not a number`;
      }

      if (args.options.limit && typeof args.options.limit !== 'number') {
        return `${args.options.limit} is not a number`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    ${chalk.yellow('Attention:')} In order to use this command, you need to grant the Azure AD
    application used by the Office 365 CLI the permission to the Yammer API.
    To do this, execute the ${chalk.blue('consent --service yammer')} command.
    
  Examples:
    
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