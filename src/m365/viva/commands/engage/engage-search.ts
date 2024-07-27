import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  queryText: string;
  show: string;
  limit?: number;
}

interface VivaEngageSearchResponse {
  count: VivaEngageSearchSummary;
  groups: VivaEngageBasicGroupResponse[];
  messages: { messages: VivaEngageBasicMessageResponse[] };
  topics: VivaEngageBasicTopicResponse[];
  users: VivaEngageBasicUserResponse[];
}

interface VivaEngageSearchSummary {
  groups: number;
  messages: number;
  topics: number;
  users: number;
}

interface VivaEngageConsolidatedResponse {
  id: string,
  description: string,
  type: string,
  web_url: string
}

interface VivaEngageBasicGroupResponse {
  id: string,
  name: string,
  web_url: string,
  state: string,
  privacy: string
  full_name: string,
  description: string,
  moderated: boolean
}

interface VivaEngageBasicTopicResponse {
  id: string,
  name: string,
  web_url: string,
  normalized_name: string,
  followers_count: number,
  description: string
}

interface VivaEngageBasicUserResponse {
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

interface VivaEngageBasicMessageResponse {
  id: string,
  created_at: string,
  group_id: string,
  thread_id: string,
  privacy: string,
  content_excerpt: string,
  web_url: string
}

class VivaEngageSearchCommand extends VivaEngageCommand {
  private static showOptions: string[] = [
    'summary', 'messages', 'users', 'topics', 'groups'
  ];

  private summary: VivaEngageSearchSummary;
  private messages: VivaEngageBasicMessageResponse[];
  private groups: VivaEngageBasicGroupResponse[];
  private topics: VivaEngageBasicTopicResponse[];
  private users: VivaEngageBasicUserResponse[];

  public get name(): string {
    return commands.ENGAGE_SEARCH;
  }

  public get description(): string {
    return 'Returns a list of messages, users, topics and groups that match the specified query.';
  }

  constructor() {
    super();
    this.summary = {
      messages: 0,
      groups: 0,
      topics: 0,
      users: 0
    };
    this.messages = [];
    this.groups = [];
    this.topics = [];
    this.users = [];

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        show: typeof args.options.show !== 'undefined',
        limit: typeof args.options.limit !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--queryText <queryText>'
      },
      {
        option: '--show [show]',
        autocomplete: VivaEngageSearchCommand.showOptions
      },
      {
        option: '--limit [limit]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.queryText && typeof args.options.queryText !== 'string') {
          return `${args.options.queryText} is not a string`;
        }

        if (args.options.limit && typeof args.options.limit !== 'number') {
          return `${args.options.limit} is not a number`;
        }

        if (args.options.output !== 'json') {
          if (typeof args.options.show !== 'undefined') {
            const scope = args.options.show.toString().toLowerCase();
            if (VivaEngageSearchCommand.showOptions.indexOf(scope) < 0) {
              return `${scope} is not a valid value for show. Allowed values are ${VivaEngageSearchCommand.showOptions.join(', ')}`;
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
    );
  }

  private async getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    const endpoint = `${this.resource}/v1/search.json?search=${formatting.encodeQueryParameter(args.options.queryText)}&page=${page}`;
    const requestOptions: CliRequestOptions = {
      url: endpoint,
      responseType: 'json',
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    const results: VivaEngageSearchResponse = await request.get<VivaEngageSearchResponse>(requestOptions);

    // results count should only read once
    if (page === 1) {
      this.summary = {
        messages: results.count.messages,
        topics: results.count.topics,
        users: results.count.users,
        groups: results.count.groups
      };
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
      await this.getAllItems(logger, args, ++page);
    }
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.summary = {
      messages: 0,
      groups: 0,
      topics: 0,
      users: 0
    };
    this.messages = [];
    this.groups = [];
    this.topics = [];
    this.users = [];

    try {
      await this.getAllItems(logger, args, 1);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(
          {
            summary: this.summary,
            messages: this.messages,
            users: this.users,
            topics: this.topics,
            groups: this.groups
          });
      }
      else {
        const show = args.options.show?.toLowerCase();

        if (show === "summary") {
          await logger.log(this.summary);
        }
        else {
          let results: VivaEngageConsolidatedResponse[] = [];
          if (show === undefined || show === "messages") {
            results = [...results, ...this.messages.map((msg) => {
              let trimmedMessage = msg.content_excerpt;
              trimmedMessage = trimmedMessage?.length >= 80 ? (trimmedMessage.substring(0, 80) + "...") : trimmedMessage;
              trimmedMessage = trimmedMessage?.replace(/\n/g, " ");
              return <VivaEngageConsolidatedResponse>
                {
                  id: msg.id,
                  description: trimmedMessage,
                  type: "message",
                  web_url: msg.web_url
                };
            })];
          }

          if (show === undefined || show === "topics") {
            results = [...results, ...this.topics.map((topic) => {
              return <VivaEngageConsolidatedResponse>
                {
                  id: topic.id,
                  description: topic.name,
                  type: "topic",
                  web_url: topic.web_url
                };
            })];
          }

          if (show === undefined || show === "users") {
            results = [...results, ...this.users.map((user) => {
              return <VivaEngageConsolidatedResponse>
                {
                  id: user.id,
                  description: user.name,
                  type: "user",
                  web_url: user.web_url
                };
            })];
          }

          if (show === undefined || show === "groups") {
            results = [...results, ...this.groups.map((group) => {
              return <VivaEngageConsolidatedResponse>
                {
                  id: group.id,
                  description: group.name,
                  type: "group",
                  web_url: group.web_url
                };
            })];
          }

          await logger.log(results);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageSearchCommand();