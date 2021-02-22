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
  queryText: string;
  show: string;
  limit?: number;
}

interface YammerSearchResponse {
  count: YammerSearchSummary;
  groups: YammerBasicGroupResponse[];
  messages: { messages: YammerBasicMessageResponse[] };
  topics: YammerBasicTopicResponse[];
  users: YammerBasicUserResponse[];
}

interface YammerSearchSummary {
  groups: number;
  messages: number;
  topics: number;
  users: number;
}

interface YammerConsolidatedResponse {
  id: string,
  description: string,
  type: string,
  web_url: string
}

interface YammerBasicGroupResponse {
  id: string,
  name: string,
  web_url: string,
  state: string,
  privacy: string
  full_name: string,
  description: string,
  moderated: boolean
}

interface YammerBasicTopicResponse {
  id: string,
  name: string,
  web_url: string,
  normalized_name: string,
  followers_count: number,
  description: string
}

interface YammerBasicUserResponse {
  id: string,
  name: string,
  web_url: string,
  state: string,
  email: string,
  first_name: string,
  last_name: string,
  full_name: string,
  admin: boolean
}

interface YammerBasicMessageResponse {
  id: string,
  created_at: string,
  group_id: string,
  thread_id: string,
  privacy: string,
  content_excerpt: string,
  web_url: string
}

class YammerSearchCommand extends YammerCommand {
  private static showOptions: string[] =
    ['summary', 'messages', 'users', 'topics', 'groups'];

  private summary: YammerSearchSummary;
  private messages: YammerBasicMessageResponse[];
  private groups: YammerBasicGroupResponse[];
  private topics: YammerBasicTopicResponse[];
  private users: YammerBasicUserResponse[];

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
    return commands.YAMMER_SEARCH;
  }

  public get description(): string {
    return 'Returns a list of messages, users, topics and groups that match the specified query.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.show = typeof args.options.show !== 'undefined';
    telemetryProps.limit = typeof args.options.limit !== 'undefined';
    return telemetryProps;
  }

  private getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const endpoint = `${this.resource}/v1/search.json?search=${encodeURIComponent(args.options.queryText)}&page=${page}`;
      const requestOptions: any = {
        url: endpoint,
        responseType: 'json'
      };

      request
        .get<YammerSearchResponse>(requestOptions)
        .then((results: YammerSearchResponse): void => {
          // results count should only read once
          if (page === 1) {
            this.summary = {
              messages: results.count.messages,
              topics: results.count.topics,
              users: results.count.users,
              groups: results.count.groups
            }
          }

          const resultMessages = results.messages.messages;
          const resultTopics = results.topics;
          const resultGroups = results.groups;
          const resultUsers = results.users;

          if (resultMessages.length > 0) {
            this.messages = this.messages.concat(resultMessages);

            if (args.options.limit && this.messages.length > args.options.limit) {
              this.messages = this.messages.slice(0, args.options.limit);
            }
          }

          if (resultTopics.length > 0) {
            this.topics = this.topics.concat(resultTopics);

            if (args.options.limit && this.topics.length > args.options.limit) {
              this.topics = this.topics.slice(0, args.options.limit);
            }
          }

          if (resultGroups.length > 0) {
            this.groups = this.groups.concat(resultGroups);

            if (args.options.limit && this.groups.length > args.options.limit) {
              this.groups = this.groups.slice(0, args.options.limit);
            }
          }

          if (resultUsers.length > 0) {
            this.users = this.users.concat(resultUsers);

            if (args.options.limit && this.users.length > args.options.limit) {
              this.users = this.users.slice(0, args.options.limit);
            }
          }

          const continueProcessing = resultMessages.length === 20 ||
            resultUsers.length === 20 ||
            resultGroups.length === 20 ||
            resultTopics.length === 20;
          if (continueProcessing) {
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
          const show = args.options.show?.toLowerCase();

          if (show === "summary") {
            logger.log(this.summary)
          }
          else {
            let results: YammerConsolidatedResponse[] = [];
            if (show === undefined || show === "messages") {
              results = [...results, ...this.messages.map((msg) => {
                let trimmedMessage = msg.content_excerpt;
                trimmedMessage = trimmedMessage?.length >= 80 ? (trimmedMessage.substring(0, 80) + "...") : trimmedMessage;
                trimmedMessage = trimmedMessage?.replace(/\n/g, " ")
                return <YammerConsolidatedResponse>
                  {
                    id: msg.id,
                    description: trimmedMessage,
                    type: "message",
                    web_url: msg.web_url
                  }
              })];
            }

            if (show === undefined || show === "topics") {
              results = [...results, ...this.topics.map((topic) => {
                return <YammerConsolidatedResponse>
                  {
                    id: topic.id,
                    description: topic.name,
                    type: "topic",
                    web_url: topic.web_url
                  }
              })];
            }

            if (show === undefined || show === "users") {
              results = [...results, ...this.users.map((user) => {
                return <YammerConsolidatedResponse>
                  {
                    id: user.id,
                    description: user.name,
                    type: "user",
                    web_url: user.web_url
                  }
              })];
            }

            if (show === undefined || show === "groups") {
              results = [...results, ...this.groups.map((group) => {
                return <YammerConsolidatedResponse>
                  {
                    id: group.id,
                    description: group.name,
                    type: "group",
                    web_url: group.web_url
                  }
              })];
            }

            logger.log(results);
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--queryText <queryText>'
      },
      {
        option: '--show [show]',
        autocomplete: YammerSearchCommand.showOptions
      },
      {
        option: '--limit [limit]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.queryText && typeof args.options.queryText !== 'string') {
      return `${args.options.queryText} is not a string`;
    }

    if (args.options.limit && typeof args.options.limit !== 'number') {
      return `${args.options.limit} is not a number`;
    }

    if (args.options.output !== 'json') {
      if (typeof args.options.show !== 'undefined') {
        const scope = args.options.show.toString().toLowerCase();
        if (YammerSearchCommand.showOptions.indexOf(scope) < 0) {
          return `${scope} is not a valid value for show. Allowed values are ${YammerSearchCommand.showOptions.join(', ')}`;
        }
      }
    }
    else {
      if (typeof args.options.show !== 'undefined') {
        return `${args.options.show} can't be used when --output set to json`;
      }
    }

    return true;
  }
}

module.exports = new YammerSearchCommand();