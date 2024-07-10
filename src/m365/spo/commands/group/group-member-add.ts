import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
  userNames?: string;
  emails?: string;
  userIds?: string;
  entraGroupIds?: string;
  aadGroupIds?: string;
  entraGroupNames?: string;
  aadGroupNames?: string;
}

class SpoGroupMemberAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_ADD;
  }

  public get description(): string {
    return 'Add members to a SharePoint Group';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'UserPrincipalName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined',
        emails: typeof args.options.emails !== 'undefined',
        userIds: typeof args.options.userIds !== 'undefined',
        entraGroupIds: typeof args.options.entraGroupIds !== 'undefined',
        aadGroupIds: typeof args.options.aadGroupIds !== 'undefined',
        entraGroupNames: typeof args.options.entraGroupNames !== 'undefined',
        aadGroupNames: typeof args.options.aadGroupNames !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--userNames [userNames]'
      },
      {
        option: '--emails [emails]'
      },
      {
        option: '--userIds [userIds]'
      },
      {
        option: '--entraGroupIds [entraGroupIds]'
      },
      {
        option: '--aadGroupIds [aadGroupIds]'
      },
      {
        option: '--entraGroupNames [entraGroupNames]'
      },
      {
        option: '--aadGroupNames [aadGroupNames]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.groupId && !validation.isValidPositiveInteger(args.options.groupId)) {
          return `Specified groupId ${args.options.groupId} is not a positive number.`;
        }

        if (args.options.userIds) {
          const isValidArray = validation.isValidPositiveIntegerArray(args.options.userIds);
          if (isValidArray !== true) {
            return `Option 'userIds' contains one or more invalid numbers: ${isValidArray}.`;
          }
        }

        if (args.options.userNames) {
          const isValidArray = validation.isValidUserPrincipalNameArray(args.options.userNames);
          if (isValidArray !== true) {
            return `Option 'userNames' contains one or more invalid UPNs: ${isValidArray}.`;
          }
        }

        if (args.options.emails) {
          const isValidArray = validation.isValidUserPrincipalNameArray(args.options.emails);
          if (isValidArray !== true) {
            return `Option 'emails' contains one or more invalid UPNs: ${isValidArray}.`;
          }
        }

        if (args.options.entraGroupIds) {
          const isValidArray = validation.isValidGuidArray(args.options.entraGroupIds);
          if (isValidArray !== true) {
            return `Option 'entraGroupIds' contains one or more invalid GUIDs: ${isValidArray}.`;
          }
        }

        if (args.options.aadGroupIds) {
          const isValidArray = validation.isValidGuidArray(args.options.aadGroupIds);
          if (isValidArray !== true) {
            return `Option 'aadGroupIds' contains one or more invalid GUIDs: ${isValidArray}.`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupName'] },
      { options: ['userNames', 'emails', 'userIds', 'entraGroupIds', 'aadGroupIds', 'entraGroupNames', 'aadGroupNames'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'groupName', 'userNames', 'emails', 'userIds', 'entraGroupIds', 'aadGroupIds', 'entraGroupNames', 'aadGroupNames');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.aadGroupIds) {
        args.options.entraGroupIds = args.options.aadGroupIds;

        await this.warn(logger, `Option 'aadGroupIds' is deprecated. Please use 'entraGroupIds' instead.`);
      }

      if (args.options.aadGroupNames) {
        args.options.entraGroupNames = args.options.aadGroupNames;

        await this.warn(logger, `Option 'aadGroupNames' is deprecated. Please use 'entraGroupNames' instead.`);
      }

      const loginNames = await this.getLoginNames(logger, args.options);

      let apiUrl = `${args.options.webUrl}/_api/web/SiteGroups`;
      if (args.options.groupId) {
        apiUrl += `/GetById(${args.options.groupId})`;
      }
      else {
        apiUrl += `/GetByName('${formatting.encodeQueryParameter(args.options.groupName!)}')`;
      }
      apiUrl += '/users';

      if (this.verbose) {
        await logger.logToStderr('Adding members to group...');
      }

      const result = [];
      for (const loginName of loginNames) {
        const requestOptions: CliRequestOptions = {
          url: apiUrl,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            LoginName: loginName
          }
        };
        const response = await request.post(requestOptions);
        result.push(response);
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getLoginNames(logger: Logger, options: Options): Promise<string[]> {
    const loginNames: string[] = [];

    if (options.userNames || options.emails) {
      loginNames.push(...formatting.splitAndTrim(options.userNames || options.emails!).map(u => `i:0#.f|membership|${u}`));
    }
    else if (options.entraGroupIds || options.entraGroupNames) {
      if (this.verbose) {
        await logger.logToStderr(`Resolving ${(options.entraGroupIds || options.entraGroupNames)!.length} group(s)...`);
      }

      const groups: Group[] = [];
      if (options.entraGroupIds) {
        const groupIds = formatting.splitAndTrim(options.entraGroupIds);

        for (const groupId of groupIds) {
          const group = await entraGroup.getGroupById(groupId);
          groups.push(group);
        }
      }
      else {
        const groupNames = formatting.splitAndTrim(options.entraGroupNames!);

        for (const groupName of groupNames) {
          const group = await entraGroup.getGroupByDisplayName(groupName);
          groups.push(group);
        }
      }

      // Check if group is M365 group or security group
      loginNames.push(...groups.map(g => g.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${g.id}` : `c:0t.c|tenant|${g.id}`));
    }
    else if (options.userIds) {
      const userIds = formatting.splitAndTrim(options.userIds);

      if (this.verbose) {
        await logger.logToStderr(`Resolving ${userIds.length} user(s)...`);
      }

      for (const userId of userIds) {
        const loginName = await this.getUserLoginNameById(options.webUrl, parseInt(userId));
        loginNames.push(loginName);
      }
    }

    return loginNames;
  }

  private async getUserLoginNameById(webUrl: string, userId: number): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/SiteUsers/GetById(${userId})?$select=LoginName`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const user = await request.get<{ LoginName: string }>(requestOptions);
    return user.LoginName;
  }
}

export default new SpoGroupMemberAddCommand();