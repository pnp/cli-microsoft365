import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand, { Options as EntraUserGetCommandOptions } from '../../../entra/commands/user/user-get.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import SpoGroupMemberListCommand, { Options as SpoGroupMemberListCommandOptions } from './group-member-list.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
  userName?: string;
  email?: string;
  userId?: number;
  entraGroupId?: string;
  aadGroupId?: string;
  entraGroupName?: string;
  aadGroupName?: string;
  force?: boolean;
}

class SpoGroupMemberRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified member from a SharePoint group';
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
        groupId: (!(!args.options.groupId)).toString(),
        groupName: (!(!args.options.groupName)).toString(),
        userName: (!(!args.options.userName)).toString(),
        email: (!(!args.options.email)).toString(),
        userId: (!(!args.options.userId)).toString(),
        entraGroupId: (!(!args.options.entraGroupId)).toString(),
        aadGroupId: (!(!args.options.groupId)).toString(),
        entraGroupName: (!(!args.options.entraGroupName)).toString(),
        aadGroupName: (!(!args.options.aadGroupName)).toString(),
        force: (!(!args.options.force)).toString()
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
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--aadGroupId [aadGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
      },
      {
        option: '--aadGroupName [aadGroupName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && isNaN(args.options.groupId)) {
          return `Specified "groupId" ${args.options.groupId} is not valid`;
        }

        if (args.options.userId && isNaN(args.options.userId)) {
          return `Specified "userId" ${args.options.userId} is not valid`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId as string)) {
          return `${args.options.entraGroupId} is not a valid GUID`;
        }

        if (args.options.aadGroupId && !validation.isValidGuid(args.options.aadGroupId as string)) {
          return `${args.options.aadGroupId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupName', 'groupId'] },
      { options: ['userName', 'email', 'userId', 'entraGroupId', 'aadGroupId', 'entraGroupName', 'aadGroupName'] }
    );
  }

  private async getUserName(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.userName) {
      return args.options.userName;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about the user ${args.options.email}`);
    }

    const options: EntraUserGetCommandOptions = {
      email: args.options.email,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    const userGetOutput: CommandOutput = await cli.executeCommandWithOutput(entraUserGetCommand as Command, { options: { ...options, _: [] } });
    const userOutput = JSON.parse(userGetOutput.stdout);
    return userOutput.userPrincipalName;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.aadGroupId) {
      args.options.entraGroupId = args.options.aadGroupId;

      this.warn(logger, `Option 'aadGroupId' is deprecated. Please use 'entraGroupId' instead`);
    }

    if (args.options.aadGroupName) {
      args.options.entraGroupName = args.options.aadGroupName;

      this.warn(logger, `Option 'aadGroupName' is deprecated. Please use 'entraGroupName' instead`);
    }

    if (args.options.force) {
      if (this.debug) {
        await logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
      }
      await this.removeUserfromSPGroup(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user ${args.options.userName || args.options.userId || args.options.email || args.options.entraGroupId || args.options.entraGroupName} from the SharePoint group?` });

      if (result) {
        await this.removeUserfromSPGroup(logger, args);
      }
    }
  }

  private async removeUserfromSPGroup(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing User ${args.options.userName || args.options.email || args.options.userId || args.options.entraGroupId || args.options.entraGroupName} from Group: ${args.options.groupId || args.options.groupName}`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${args.options.groupId}')`
      : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}`;

    if (args.options.userId) {
      requestUrl += `/users/removeById(${args.options.userId})`;
    }
    else if (args.options.userName || args.options.email) {
      const userName: string = await this.getUserName(logger, args);
      const loginName: string = `i:0#.f|membership|${userName}`;
      requestUrl += `/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`;
    }
    else {
      const entraGroupId = await this.getGroupId(args);
      requestUrl += `/users/RemoveById(${entraGroupId})`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const options: SpoGroupMemberListCommandOptions = {
      webUrl: args.options.webUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    if (args.options.groupId) {
      options.groupId = args.options.groupId;
    }
    else {
      options.groupName = args.options.groupName;
    }

    const output = await cli.executeCommandWithOutput(SpoGroupMemberListCommand as Command, { options: { ...options, _: [] } });
    const getGroupMemberListOutput = JSON.parse(output.stdout);

    let foundGroups: any;

    if (args.options.entraGroupId) {
      foundGroups = getGroupMemberListOutput.filter((x: any) => { return x.LoginName.indexOf(args.options.entraGroupId) > -1 && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }
    else {
      foundGroups = getGroupMemberListOutput.filter((x: any) => { return x.Title === args.options.entraGroupName && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }

    if (foundGroups.length === 0) {
      throw `The Azure AD group ${args.options.entraGroupId || args.options.entraGroupName} is not found in SharePoint group ${args.options.groupId || args.options.groupName}`;
    }

    return foundGroups[0].Id;
  }
}

export default new SpoGroupMemberRemoveCommand();