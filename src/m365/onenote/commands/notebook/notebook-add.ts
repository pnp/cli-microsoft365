import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { AxiosRequestConfig } from 'axios';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  webUrl?: string;
}

class OneNoteNotebookAddCommand extends GraphCommand {
  public get name(): string {
    return commands.NOTEBOOK_ADD;
  }

  public get description(): string {
    return 'Create a new OneNote notebook.';
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
        joined: args.options.joined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-n, --name <name>' },
      { option: '--userId [userId]' },
      { option: '--userName [userName]' },
      { option: '--groupId [groupId]' },
      { option: '--groupName [groupName]' },
      { option: '-u, --webUrl [webUrl]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.userId && args.options.userName) {
          return 'Specify either userId or userName, but not both';
        }

        if (args.options.groupId && args.options.groupName) {
          return 'Specify either groupId or groupName, but not both';
        }

        return true;
      }
    );
  }

  private async getEndpointUrl(args: CommandArgs): Promise<string> {
    let endpoint: string = `${this.resource}/v1.0/`;

    if (args.options.userId) {
      endpoint += `users/${args.options.userId}`;
    }
    else if (args.options.userName) {
      endpoint += `users/${args.options.userName}`;
    }
    else if (args.options.groupId) {
      endpoint += `groups/${args.options.groupId}`;
    }
    else if (args.options.groupName) {
      const groupId = await this.getGroupId(args.options.groupName);
      endpoint += `groups/${groupId}`;
    }
    else if (args.options.webUrl) {
      const siteId = await this.getSpoSiteId(args.options.webUrl);
      endpoint += `sites/${siteId}`;
    }
    else {
      endpoint += 'me';
    }
    endpoint += '/onenote/notebooks';
    return endpoint;
  }

  public defaultProperties(): string[] | undefined {
    return ['createdDateTime', 'displayName', 'id'];
  }

  private async getGroupId(groupName: string): Promise<string> {
    const group = await aadGroup.getGroupByDisplayName(groupName);
    return group.id!;
  }

  private async getSpoSiteId(webUrl: string): Promise<string> {
    const url = new URL(webUrl);
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${url.hostname}:${url.pathname}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const site = await request.get<{ id: string }>(requestOptions);
    return site.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const endpoint = await this.getEndpointUrl(args);
      const requestBody = {
        displayName: args.options.name
      };
      const requestOptionsPost: AxiosRequestConfig = {
        url: endpoint,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: requestBody
      };
      const response = await request.post(requestOptionsPost);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new OneNoteNotebookAddCommand();