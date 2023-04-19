import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { aadUser } from '../../../../utils/aadUser';
import { SharingResult } from './SharingResult';
import { aadGroup } from '../../../../utils/aadGroup';
import { spo } from '../../../../utils/spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
  userName?: string;
  email?: string;
  userId?: string;
  aadGroupId?: string;
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
        userName: typeof args.options.userName !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        aadGroupId: typeof args.options.aadGroupId !== 'undefined',
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
        option: '--userName [userName]'
      },
      {
        option: '--email [email]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--aadGroupId [aadGroupId]'
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
        if (args.options.userId && !userIdReg.test(args.options.userId!)) {
          return `${args.options.userId} is not a number or a comma seperated value`;
        }

        if (args.options.userName && args.options.userName.split(',').some(e => !validation.isValidUserPrincipalName(e))) {
          return `${args.options.userName} contains one or more invalid usernames`;
        }

        if (args.options.email && args.options.email.split(',').some(e => !validation.isValidUserPrincipalName(e))) {
          return `${args.options.email} contains one or more invalid email addresses`;
        }

        if (args.options.aadGroupId && args.options.aadGroupId.split(',').some(e => !validation.isValidGuid(e))) {
          return `${args.options.aadGroupId} contains one or more invalid GUIDs`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupName'] },
      { options: ['userName', 'email', 'userId', 'aadGroupId', 'aadGroupName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args, logger);
      const resolvedUsernameList = await this.getValidUsers(args, logger);

      if (this.verbose) {
        logger.logToStderr(`Adding resource(s) to SharePoint Group ${args.options.groupId || args.options.groupName}`);
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

      logger.log(sharingResult.UsersAddedToGroup);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getGroupId(args: CommandArgs, logger: Logger): Promise<number> {
    if (this.verbose) {
      logger.logToStderr(`Getting group Id for SharePoint Group ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const getGroupMethod: string = args.options.groupName ?
      `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')` :
      `GetById('${args.options.groupId}')`;

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/sitegroups/${getGroupMethod}?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ Id: number }>(requestOptions)
      .then(response => response.Id);
  }

  private async getValidUsers(args: CommandArgs, logger: Logger): Promise<string[]> {
    if (this.verbose) {
      logger.logToStderr('Checking if the specified users and groups exist');
    }

    const validUserNames: string[] = [];
    const identifiers: string = args.options.userName ?? args.options.email ?? args.options.aadGroupId ?? args.options.aadGroupName ?? args.options.userId!.toString();

    await Promise.all(identifiers.split(',').map(async identifier => {
      const trimmedIdentifier = identifier.trim();
      try {
        if (args.options.userId) {
          if (this.verbose) {
            logger.logToStderr(`Getting AAD ID of user with ID ${trimmedIdentifier}`);
          }
          const spoUserAzureId = await spo.getUserAzureIdBySpoId(args.options.webUrl, trimmedIdentifier);
          validUserNames.push(spoUserAzureId);
        }
        else if (args.options.userName) {
          validUserNames.push(trimmedIdentifier);
        }
        else if (args.options.aadGroupId) {
          validUserNames.push(trimmedIdentifier);
        }
        else if (args.options.aadGroupName) {
          if (this.verbose) {
            logger.logToStderr(`Getting ID of Azure AD group ${trimmedIdentifier}`);
          }
          const groupId = await aadGroup.getGroupIdByDisplayName(trimmedIdentifier);
          validUserNames.push(groupId);
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Getting Azure AD ID for user ${trimmedIdentifier}`);
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

module.exports = new SpoGroupMemberAddCommand();