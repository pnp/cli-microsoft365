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
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Options as SpoGroupMemberListCommandOptions } from './group-member-list';
import * as SpoGroupMemberListCommand from './group-member-list';

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
  aadGroupId?: string;
  aadGroupName?: string;
  confirm?: boolean;
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
        confirm: (!(!args.options.confirm)).toString()
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
      },
      {
        option: '--confirm'
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
      { options: ['userName', 'email', 'userId', 'aadGroupId', 'aadGroupName'] }
    );
  }

  private async getUserName(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.userName) {
      return args.options.userName;
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about the user ${args.options.email}`);
    }

    const options: AadUserGetCommandOptions = {
      email: args.options.email,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    const userGetOutput: CommandOutput = await Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } });
    const userOutput = JSON.parse(userGetOutput.stdout);
    return userOutput.userPrincipalName;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
      }
      await this.removeUserfromSPGroup(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove user ${args.options.userName || args.options.userId || args.options.email || args.options.aadGroupId || args.options.aadGroupName} from the SharePoint group?`
      });

      if (result.continue) {
        await this.removeUserfromSPGroup(logger, args);
      }
    }
  }

  private async removeUserfromSPGroup(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing User ${args.options.userName || args.options.email || args.options.userId || args.options.aadGroupId || args.options.aadGroupName} from Group: ${args.options.groupId || args.options.groupName}`);
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
      const aadGroupId = await this.getGroupId(args);
      requestUrl += `/users/RemoveById(${aadGroupId})`;
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

    const output = await Cli.executeCommandWithOutput(SpoGroupMemberListCommand as Command, { options: { ...options, _: [] } });
    const getGroupMemberListOutput = JSON.parse(output.stdout);

    let foundgroups: any;

    if (args.options.aadGroupId) {
      foundgroups = getGroupMemberListOutput.filter((x: any) => { return x.LoginName.indexOf(args.options.aadGroupId) > -1 && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }
    else {
      foundgroups = getGroupMemberListOutput.filter((x: any) => { return x.Title === args.options.aadGroupName && (x.LoginName.indexOf("c:0o.c|federateddirectoryclaimprovider|") === 0 || x.LoginName.indexOf("c:0t.c|tenant|") === 0); });
    }

    if (foundgroups.length === 0) {
      throw `The Azure AD group ${args.options.aadGroupId || args.options.aadGroupName} is not found in SharePoint group ${args.options.groupId || args.options.groupName}`;
    }

    return foundgroups[0].Id;
  }
}

module.exports = new SpoGroupMemberRemoveCommand();