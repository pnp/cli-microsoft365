import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import * as SpoUserGetCommand from '../user/user-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SharingResult } from './SharingResult';

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
        userId: typeof args.options.userId !== 'undefined'
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupName'] },
      { options: ['userName', 'email', 'userId'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args, logger);
      const resolvedUsernameList = await this.getValidUsers(args, logger);

      if (this.verbose) {
        logger.logToStderr(`Adding user(s) to SharePoint Group ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
      }

      const data: any = {
        url: args.options.webUrl,
        peoplePickerInput: this.getFormattedUserList(resolvedUsernameList),
        roleValue: `group:${groupId}`
      };

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/SP.Web.ShareObject`,
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose'
        },
        data: data,
        responseType: 'json'
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
      url: `${args.options.webUrl}/_api/web/sitegroups/${getGroupMethod}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ Id: number }>(requestOptions)
      .then(response => {
        const groupId: number | undefined = response.Id;

        if (!groupId) {
          return Promise.reject(`The specified group does not exist in the SharePoint site`);
        }

        return groupId;
      });
  }

  private getValidUsers(args: CommandArgs, logger: Logger): Promise<string[]> {
    if (this.verbose) {
      logger.logToStderr(`Checking if the specified users exist`);
    }

    const validUserNames: string[] = [];
    const invalidUserNames: string[] = [];
    const userIdentifiers: string = args.options.userName || args.options.email || args.options.userId!.toString();

    return Promise
      .all(userIdentifiers.split(',').map(async userIdentifier => {
        try {
          if (args.options.userId) {
            await this.spoUserGet(args.options, userIdentifier.trim(), logger, validUserNames);
          }
          else {
            await this.aadUserGet(args.options, userIdentifier.trim(), logger, validUserNames);
          }
        }
        catch (err: any) {
          logger.logToStderr(err.stderr);
          invalidUserNames.push(userIdentifier);

          return err;
        }
      }))
      .then((): Promise<string[]> => {
        if (invalidUserNames.length > 0) {
          return Promise.reject(`Users not added to the group because the following users don't exist: ${invalidUserNames.join(', ')}`);
        }

        return Promise.resolve(validUserNames);
      });
  }

  private async aadUserGet(options: Options, userIdentifier: string, logger: Logger, validUserNames: string[]): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Get UPN from Azure AD for user ${userIdentifier}`);
    }

    const aadUserGetCommandoptions: AadUserGetCommandOptions = {
      ...(options.userName && { userName: userIdentifier }),
      ...(options.email && { email: userIdentifier }),
      output: 'json',
      debug: options.debug,
      verbose: options.verbose
    };

    const aadUserGetOutput: CommandOutput = await Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...aadUserGetCommandoptions, _: [] } });

    if (this.debug) {
      logger.logToStderr(aadUserGetOutput.stderr);
    }

    validUserNames.push(JSON.parse(aadUserGetOutput.stdout).userPrincipalName);
  }

  private async spoUserGet(options: Options, userIdentifier: string, logger: Logger, validUserNames: string[]): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Get UPN from SharePoint for user ${userIdentifier}`);
    }

    const spoUserGetCommandoptions: SpoUserGetCommandOptions = {
      id: userIdentifier,
      webUrl: options.webUrl,
      output: 'json',
      debug: options.debug,
      verbose: options.verbose
    };

    const spoUserGetOutput: CommandOutput = await Cli.executeCommandWithOutput(SpoUserGetCommand as Command, { options: { ...spoUserGetCommandoptions, _: [] } });

    if (this.debug) {
      logger.logToStderr(spoUserGetOutput.stderr);
    }

    validUserNames.push(JSON.parse(spoUserGetOutput.stdout).UserPrincipalName);
  }

  private getFormattedUserList(activeUserList: string[]): any {
    const generatedPeoplePicker: any = JSON.stringify(activeUserList.map(singleUsername => {
      return { Key: singleUsername.trim() };
    }));

    return generatedPeoplePicker;
  }
}

module.exports = new SpoGroupMemberAddCommand();