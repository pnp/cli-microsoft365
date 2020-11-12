import { Logger } from '../../../cli';
import { CommandOption } from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import YammerCommand from '../../base/YammerCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  search: string;
  limit?: number;
}

interface YammerSearchResponse {
  count: {
    groups: number;
    messages: number;
    topics: number;
    users: number;
  };
  groups: any[];
  messages: { messages: any[] };
  topics: any[];
  users: any[];
}

class YammerSearchCommand extends YammerCommand {
  private summary: any;
  private messages: any[];
  private groups: any[];
  private topics: any[];
  private users: any[];

  constructor() {
    super();
    this.summary = {
      messages: 0,
      groups: 0,
      topics: 0,
      users: 0
    }
    this.messages = [];
    this.groups = [];
    this.topics = [];
    this.users = [];
  }

  public get name(): string {
    return `${commands.YAMMER_SEARCH}`;
  }

  public get description(): string {
    return 'Returns a list of messages, users, topics and groups that match the user’s search query.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.limit =  typeof args.options.limit !== 'undefined';
    return telemetryProps;
  }

  private getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const endpoint = `${this.resource}/v1/search.json?search=${args.options.search}&page=${page}`;
      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get<YammerSearchResponse>(requestOptions)
        .then((results: YammerSearchResponse): void => {
          // results count should only read once
          if (page === 1) {
            const summary = results.count;
            this.summary = {
              messages: summary.messages,
              topics: summary.topics,
              users: summary.users,
              groups: summary.groups
            }
          }

          const resultMessages = results.messages.messages;
          const resultTopics = results.topics;
          const resultGroups = results.groups;
          const resultUsers = results.users;

          let continueProcessing    = true;
          let resultsFound = false;

          if (resultMessages.length > 0) {
            resultsFound = true;
            this.messages = this.messages.concat(resultMessages);

            if (args.options.limit && this.messages.length > args.options.limit) {
              continueProcessing    = false;
              this.messages = this.messages.slice(0, args.options.limit);
            }
          }

          if (resultTopics.length > 0) {
            resultsFound = true;
            this.topics = this.topics.concat(resultTopics);

            if (args.options.limit && this.topics.length > args.options.limit) {
              continueProcessing    = false;
              this.topics = this.topics.slice(0, args.options.limit);
            }
          }

          if (resultGroups.length > 0) {
            resultsFound = true;
            this.groups = this.groups.concat(resultGroups);

            if (args.options.limit && this.groups.length > args.options.limit) {
              continueProcessing    = false;
              this.groups = this.groups.slice(0, args.options.limit);
            }
          }

          if (resultUsers.length > 0) {
            resultsFound = true;
            this.users = this.users.concat(resultUsers);

            if (args.options.limit && this.users.length > args.options.limit) {
              continueProcessing    = false;
              this.users = this.users.slice(0, args.options.limit);
            }
          }

          // does not process more queries if we are not able to return a complex object 
          if (resultsFound && continueProcessing && args.options.output === 'json') {
            this
                .getAllItems(logger, args, ++page)
                .then((): void => {
                  resolve();
                }, (err: any): void => {
                  reject(err);
                });
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  };

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.summary = {
      messages: 0,
      groups: 0,
      topics: 0,
      users: 0
    }
    this.messages = [];
    this.groups = [];
    this.topics = [];
    this.users = [];

    this
      .getAllItems(logger, args, 1)
      .then((): void => {
        if (args.options.output === 'json') {
          logger.log(
            {
              summary: this.summary,
              messages: this.messages,
              users: this.users,
              topics: this.topics,
              groups: this.groups,
            });
        }
        else {
          logger.log(this.summary);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --search <search>',
        description: 'The query for the search'
      },
      {
        option: '--limit [limit]',
        description: 'Limits the results returned for each item category. Can only be used with the --output json option.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.search && typeof args.options.search !== 'string') {
      return `${args.options.search} is not a string`;
    }

    if (args.options.limit && args.options.output !== 'json') {
      return 'Limit can only used with the parameter --output json'
    }

    if (args.options.limit && typeof args.options.limit !== 'number') {
      return `${args.options.limit} is not a number`;
    }

    return true;
  }
}

module.exports = new YammerSearchCommand();