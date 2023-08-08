import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SharingResult } from './SharingResult.js';

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
  aadGroupIds?: string;
  aadGroupName?: string;
}

class SpoGroupMemberAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_ADD;
  }

  public get description(): string {
    return 'Add members to a SharePoint Group';
  }

  public defaultProperties(): string[] | undefined {
    return ['DisplayName', 'Email'];
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined',
        emails: typeof args.options.emails !== 'undefined',
        userIds: typeof args.options.userIds !== 'undefined',
        aadGroupIds: typeof args.options.aadGroupIds !== 'undefined',
        aadGroupName: typeof args.options.aadGroupName !== 'undefined'
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
        option: '--aadGroupIds [aadGroupIds]'
      },
      {
        option: '--aadGroupName [aadGroupName]'
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

        if (args.options.groupId && isNaN(args.options.groupId)) {
          return `Specified groupId ${args.options.groupId} is not a number`;
        }

        const userIdReg = new RegExp(/^[0-9,]*$/);
        if (args.options.userIds && !userIdReg.test(args.options.userIds!)) {
          return `${args.options.userIds} is not a number or a comma seperated value`;
        }

        if (args.options.userNames && args.options.userNames.split(',').some(e => !validation.isValidUserPrincipalName(e))) {
          return `${args.options.userNames} contains one or more invalid usernames`;
        }

        if (args.options.emails && args.options.emails.split(',').some(e => !validation.isValidUserPrincipalName(e))) {
          return `${args.options.emails} contains one or more invalid email addresses`;
        }

        if (args.options.aadGroupIds && args.options.aadGroupIds.split(',').some(e => !validation.isValidGuid(e))) {
          return `${args.options.aadGroupIds} contains one or more invalid GUIDs`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupName'] },
      { options: ['userNames', 'emails', 'userIds', 'aadGroupIds', 'aadGroupName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args, logger);
      const resolvedUsernameList = await this.getValidUsers(args, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding resource(s) to SharePoint Group ${args.options.groupId || args.options.groupName}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/SP.Web.ShareObject`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          url: args.options.webUrl,
          peoplePickerInput: this.getFormattedUserList(resolvedUsernameList),
          roleValue: `group:${groupId}`
        }
      };

      const sharingResult = await request.post<SharingResult>(requestOptions);
      if (sharingResult.ErrorMessage !== null) {
        throw sharingResult.ErrorMessage;
      }

      await logger.log(sharingResult.UsersAddedToGroup);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs, logger: Logger): Promise<number> {
    if (this.verbose) {
      await logger.logToStderr(`Getting group Id for SharePoint Group ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const getGroupMethod: string = args.options.groupName ?
      `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')` :
      `GetById('${args.options.groupId}')`;

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/sitegroups/${getGroupMethod}?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ Id: number }>(requestOptions);
    return response.Id;
  }

  private async getValidUsers(args: CommandArgs, logger: Logger): Promise<string[]> {
    if (this.verbose) {
      await logger.logToStderr('Checking if the specified users and groups exist');
    }

    const validUserNames: string[] = [];
    const identifiers: string = args.options.userNames ?? args.options.emails ?? args.options.aadGroupIds ?? args.options.aadGroupName ?? args.options.userIds!.toString();

    await Promise.all(identifiers.split(',').map(async identifier => {
      const trimmedIdentifier = identifier.trim();
      try {
        if (args.options.userIds) {
          if (this.verbose) {
            await logger.logToStderr(`Getting AAD ID of user with ID ${trimmedIdentifier}`);
          }
          const spoUserAzureId = await spo.getUserAzureIdBySpoId(args.options.webUrl, trimmedIdentifier);
          validUserNames.push(spoUserAzureId);
        }
        else if (args.options.userNames) {
          validUserNames.push(trimmedIdentifier);
        }
        else if (args.options.aadGroupIds) {
          validUserNames.push(trimmedIdentifier);
        }
        else if (args.options.aadGroupName) {
          if (this.verbose) {
            await logger.logToStderr(`Getting ID of Azure AD group ${trimmedIdentifier}`);
          }
          const groupId = await aadGroup.getGroupIdByDisplayName(trimmedIdentifier);
          validUserNames.push(groupId);
        }
        else {
          if (this.verbose) {
            await logger.logToStderr(`Getting Azure AD ID for user ${trimmedIdentifier}`);
          }
          const upn = await aadUser.getUserIdByEmail(trimmedIdentifier);
          validUserNames.push(upn);
        }
      }
      catch (err: any) {
        throw `Resource '${trimmedIdentifier}' does not exist.`;
      }
    }));

    return validUserNames;
  }

  private getFormattedUserList(activeUserList: string[]): any {
    const generatedPeoplePicker: any = JSON.stringify(activeUserList.map(singleUsername => {
      return { Key: singleUsername.trim() };
    }));

    return generatedPeoplePicker;
  }
}

export default new SpoGroupMemberAddCommand();