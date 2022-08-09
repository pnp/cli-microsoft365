import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  olderThanId?: number;
  threaded?: boolean;
  limit?: number;
  feedType?: string;
  groupId?: number;
  threadId?: number;
}

class YammerMessageListCommand extends YammerCommand {
  private items: any[];
  private static readonly feedTypes: string[] = ['All', 'Top', 'My', 'Following', 'Sent', 'Private', 'Received'];

  public get name(): string {
    return commands.MESSAGE_LIST;
  }

  public get description(): string {
    return 'Returns all accessible messages from the user\'s Yammer network';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'replied_to_id', 'thread_id', 'group_id', 'shortBody'];
  }

  constructor() {
    super();
    this.items = [];

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        olderThanId: args.options.olderThanId !== undefined,
        threaded: args.options.threaded,
        limit: args.options.limit !== undefined,
        feedType: args.options.feedType !== undefined,
        threadId: args.options.threadId !== undefined,
        groupId: args.options.groupId !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--olderThanId [olderThanId]'
      },
      {
        option: '-f, --feedType [feedType]',
        autocomplete: YammerMessageListCommand.feedTypes
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--threadId [threadId]'
      },
      {
        option: '--threaded'
      },
      {
        option: '--limit [limit]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && args.options.threadId) {
          return `You cannot specify groupId and threadId at the same time`;
        }

        if (args.options.feedType && (args.options.groupId || args.options.threadId)) {
          return `You cannot specify the feedType with groupId or threadId at the same time`;
        }

        if (args.options.feedType && YammerMessageListCommand.feedTypes.indexOf(args.options.feedType) < 0) {
          return `${args.options.feedType} is not a valid value for the feedType option. Allowed values are ${YammerMessageListCommand.feedTypes.join('|')}`;
        }

        if (args.options.olderThanId && typeof args.options.olderThanId !== 'number') {
          return `${args.options.olderThanId} is not a number`;
        }

        if (args.options.groupId && typeof args.options.groupId !== 'number') {
          return `${args.options.groupId} is not a number`;
        }

        if (args.options.threadId && typeof args.options.threadId !== 'number') {
          return `${args.options.threadId} is not a number`;
        }

        if (args.options.limit && typeof args.options.limit !== 'number') {
          return `${args.options.limit} is not a number`;
        }

        return true;
      }
    );
  }

  private getAllItems(logger: Logger, args: CommandArgs, messageId: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      let endpoint = `${this.resource}/v1`;

      if (args.options.threadId) {
        endpoint += `/messages/in_thread/${args.options.threadId}.json`;
      }
      else if (args.options.groupId) {
        endpoint += `/messages/in_group/${args.options.groupId}.json`;
      }
      else {
        if (!args.options.feedType) {
          args.options.feedType = "All";
        }

        switch (args.options.feedType) {
          case 'Top':
            endpoint += `/messages/algo.json`;
            break;
          case 'My':
            endpoint += `/messages/my_feed.json`;
            break;
          case 'Following':
            endpoint += `/messages/following.json`;
            break;
          case 'Sent':
            endpoint += `/messages/sent.json`;
            break;
          case 'Private':
            endpoint += `/messages/private.json`;
            break;
          case 'Received':
            endpoint += `/messages/received.json`;
            break;
          default:
            endpoint += `/messages.json`;
        }
      }

      if (messageId !== -1) {
        endpoint += `?older_than=${messageId}`;
      }
      else if (args.options.olderThanId) {
        endpoint += `?older_than=${args.options.olderThanId}`;
      }

      if (args.options.threaded) {
        if (endpoint.indexOf("?") > -1) {
          endpoint += "&";
        }
        else {
          endpoint += "?";
        }

        endpoint += `threaded=true`;
      }

      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
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
              this
                .getAllItems(logger, args, this.items[this.items.length - 1].id)
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
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.items = []; // this will reset the items array in interactive mode

    this
      .getAllItems(logger, args, -1)
      .then((): void => {
        this.items.forEach(m => {
          let shortBody;
          const bodyToProcess = m.body.plain;

          if (bodyToProcess) {
            let maxLength = 35;
            let addedDots = "...";
            if (bodyToProcess.length < maxLength) {
              maxLength = bodyToProcess.length;
              addedDots = "";
            }

            shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
          }

          m.shortBody = shortBody;
        });

        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new YammerMessageListCommand();