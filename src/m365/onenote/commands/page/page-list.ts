import { OnenotePage } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  webUrl?: string;
}

class OneNotePageListCommand extends GraphCommand {
  public get name(): string {
    return commands.PAGE_LIST;
  }

  public get description(): string {
    return 'Retrieve a list of OneNote pages.';
  }

  public defaultProperties(): string[] | undefined {
    return ['createdDateTime', 'title', 'id'];
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
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
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
      const siteId = await spo.getSpoGraphSiteId(args.options.webUrl);
      endpoint += `sites/${siteId}`;
    }
    else {
      endpoint += 'me';
    }
    endpoint += '/onenote/pages';
    return endpoint;
  }

  private async getGroupId(groupName: string): Promise<string> {
    const group = await aadGroup.getGroupByDisplayName(groupName);
    return group.id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const endpoint = await this.getEndpointUrl(args);
      const items = await odata.getAllItems<OnenotePage>(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OneNotePageListCommand();