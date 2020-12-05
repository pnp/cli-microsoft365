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

interface YammerBasicSearchResponse {
  id: string,
  name: string,
  url: string
  type: string
}

interface YammerBasicGroupResponse {
  id: string,
  name: string,
  url: string,
  state: string,
  privacy: string
  full_name: string,
  description: string,
  moderated: boolean
}

interface YammerBasicTopicResponse {
  id: string,
  name: string,
  url: string,
  normalized_name: string,
  followers_count: number,
  description: string
}

interface YammerBasicUserResponse {
  id: string,
  name: string,
  url: string,
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
  url: string
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
    return `${commands.YAMMER_SEARCH}`;
  }

  public get description(): string {
    return 'Returns a list of messages, users, topics and groups that match the user’s search query.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.show =  typeof args.options.show !== 'undefined';
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

          const continueProcessing = resultMessages.length === 20 || resultUsers.length === 20 || resultGroups.length === 20 || resultTopics.length === 20;
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
          if (show === "messages") {
            logger.log(this.messages.map((message) => {
              return <YammerBasicMessageResponse>
              {
                id: message.id,
                content_excerpt: encodeURI(message.content_excerpt),
                created_at: message.created_at,
                group_id: message.group_id,
                thread_id: message.thread_id,
                privacy: message.privacy,
                url: message.url
              }
            }));
          } else if (show === "users") {
            logger.log(this.users.map((user) => {
              return <YammerBasicUserResponse>
              {
                id: user.id,
                name: user.name,
                first_name: user.first_name,
                last_name: user.last_name,
                full_name: user.full_name,
                email: user.email,
                admin: user.admin,
                state: user.state,
                url: user.url
              }
            }));
          } else if (show === "topics") {
            logger.log(this.topics.map((topic) => {
              return <YammerBasicTopicResponse>
              {
                id: topic.id,
                name: topic.name,
                normalized_name: topic.normalized_name,
                description: topic.description,
                followers_count: topic.followers_count,
                url: topic.url
              }
            }));
          } else if (show === "groups") {
            logger.log(this.groups.map((group) => {
              return <YammerBasicGroupResponse>
              {
                id: group.id,
                name: group.name,
                full_name: group.full_name,
                description: group.description,
                privacy: group.privacy,
                moderated: group.moderated,
                state: group.state,
                url: group.url
              }
            }));
          } else if (show === "summary") {
            logger.log(this.summary)
          } else { 
            let results: YammerBasicSearchResponse[] = [];
            results = [...results, ...this.messages.map((msg) => {
              return <YammerBasicSearchResponse>
              {
                id: msg.id,
                name: encodeURI(msg.content_excerpt),
                url: msg.url,
                type: "message"
              }
            })]
            results = [...results, ...this.topics.map((topic) => {
              return <YammerBasicSearchResponse>
              {
                id: topic.id,
                name: topic.name,
                url: topic.url,
                type: "topic"
              }
            })]
            results = [...results, ...this.users.map((user) => {
              return <YammerBasicSearchResponse>
              {
                id: user.id,
                name: user.name,
                url: user.url,
                type: "user"
              }
            })]
            results = [...results, ...this.groups.map((group) => {
              return <YammerBasicSearchResponse>
              {
                id: group.id,
                name: group.name,
                url: group.url,
                type: "group"
              }
            })];

            logger.log(results);
          }
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
        option: '--show [show]',
        description: 'Specifies the type of data to return when using --ouptut text. Allowed values Summary|Messages|Users|Topics|Groups. Defaults to Summary.'
      },
      {
        option: '--limit [limit]',
        description: 'Limits the results returned for each item category.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.search && typeof args.options.search !== 'string') {
      return `${args.options.search} is not a string`;
    }

    if (args.options.limit && typeof args.options.limit !== 'number') {
      return `${args.options.limit} is not a number`;
    }

    if (args.options.output !== 'json') {
      if (typeof args.options.show !== 'undefined') {
        const scope = args.options.show.toString().toLowerCase();
        if (YammerSearchCommand.showOptions.indexOf(scope) < 0) {
          return `${args.options.scope} is not a valid value for show. Allowed values are summary|messages|users|topics|groups`;
        }
      }
    }
    else {
      if (typeof args.options.show !== 'undefined') {
        return `${args.options.scope} can't be used in --output json`;
      }
    }

    return true;
  }
}

module.exports = new YammerSearchCommand();