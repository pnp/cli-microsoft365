import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';

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
    return 'Create a new OneNote notebook';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        webUrl: typeof args.options.webUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '-u, --webUrl [webUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        // check name for invalid characters
        if (args.options.name.length > 128) {
          return 'The specified name is too long. It should be less than 128 characters';
        }

        if (/[?*/:<>|'"]/.test(args.options.name)) {
          return `The specified name contains invalid characters. It cannot contain ?*/:<>|'". Please remove them and try again.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.webUrl) {
          return validation.isValidSharePointUrl(args.options.webUrl);
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName', 'groupId', 'groupName', 'webUrl'],
      runsWhen: (args) => {
        const options = [args.options.userId, args.options.userName, args.options.groupId, args.options.groupName, args.options.webUrl];
        return options.some(item => item !== undefined);
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Creating OneNote notebook ${args.options.name}`);
      }

      const requestUrl = await this.getRequestUrl(args);
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': "application/json"
        },
        responseType: 'json',
        data: {
          displayName: args.options.name
        }
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getRequestUrl(args: CommandArgs): Promise<string> {
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
      const groupId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
      endpoint += `groups/${groupId}`;
    }
    else if (args.options.webUrl) {
      const siteId = await spo.getSpoGraphSiteId(args.options.webUrl);
      endpoint += `sites/${siteId}`;
    }
    else {
      endpoint += 'me';
    }
    endpoint += '/onenote/notebooks';
    return endpoint;
  }
}

export default new OneNoteNotebookAddCommand();