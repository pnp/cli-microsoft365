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

  private getEndpointUrl(args: CommandArgs): Promise<string> {
    return new Promise<string>((resolve: (endpoint: string) => void, reject: (error: string) => void): void => {
      let endpoint: string = `${this.resource}/v1.0/me/onenote/notebooks`;

      if (args.options.userId) {
        endpoint = `${this.resource}/v1.0/users/${args.options.userId}/onenote/notebooks`;
        return resolve(endpoint);
      }
      else if (args.options.userName) {
        endpoint = `${this.resource}/v1.0/users/${args.options.userName}/onenote/notebooks`;
        return resolve(endpoint);
      }
      else if (args.options.groupId) {
        endpoint = `${this.resource}/v1.0/groups/${args.options.groupId}/onenote/notebooks`;
        return resolve(endpoint);
      }
      else if (args.options.groupName) {
        this
          .getGroupId(args)
          .then((retrievedgroupId: string): void => {
            endpoint = `${this.resource}/v1.0/groups/${retrievedgroupId}/onenote/notebooks`;
            return resolve(endpoint);
          })
          .catch((err: any) => {
            reject(err);
          });
      }
      else if (args.options.webUrl) {
        this
          .getSpoSiteId(args)
          .then((siteId: string): void => {
            endpoint = `${this.resource}/v1.0/sites/${siteId}/onenote/notebooks`;
            return resolve(endpoint);
          })
          .catch((err: any) => {
            reject(err);
          });
      }
      else {
        return resolve(endpoint);
      }
    });
  }

  public defaultProperties(): string[] | undefined {
    return ['createdDateTime', 'displayName', 'id'];
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    return aadGroup
      .getGroupByDisplayName(args.options.groupName!)
      .then(group => group.id!);
  }

  private getSpoSiteId(args: CommandArgs): Promise<string> {
    const url = new URL(args.options.webUrl!);
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${url.hostname}:${url.pathname}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ id: string }>(requestOptions)
      .then((site: { id: string }) => site.id);
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