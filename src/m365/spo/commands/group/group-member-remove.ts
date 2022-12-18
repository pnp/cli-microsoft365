import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Options as SpoGroupMemberListCommandOptions } from './group-member-list';
import * as SpoGroupMemberListCommand from './group-member-list';
import Command from '../../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
  userName?: string;
  confirm?: boolean;
  aadGroupId?: string;
  aadGroupName?: string;
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        aadGroupId: typeof args.options.aadGroupId !== 'undefined',
        aadGroupName: typeof args.options.aadGroupName !== 'undefined',
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
        if (args.options.groupId && isNaN(args.options.groupId)) {
          return `Specified "groupId" ${args.options.groupId} is not valid`;
        }

        if (args.options.aadGroupId && !validation.isValidGuid(args.options.aadGroupId as string)) {
          return `${args.options.aadGroupId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupName', 'groupId'] });
    this.optionSets.push({ options: ['userName', 'aadGroupId', 'aadGroupName'] });
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
        message: `Are you sure you want to remove user User ${args.options.userName} from SharePoint group?`
      });

      if (result.continue) {
        await this.removeUserfromSPGroup(logger, args);
      }
    }
  }

  private async removeUserfromSPGroup(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing User with Username ${args.options.userName} from Group: ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${args.options.groupId}')`
      : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}`;

    if (args.options.userName) {
      const loginName: string = `i:0#.f|membership|${args.options.userName}`;
      requestUrl += `/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`;
    }
    else {
      const aadGroupId = await this.getGroupId(args);
      requestUrl += `/users/RemoveById(${aadGroupId})`;
      logger.log(aadGroupId);
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