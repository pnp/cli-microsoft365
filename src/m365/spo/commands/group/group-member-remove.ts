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

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupName', 'groupId'] },
      { options: ['userName', 'email', 'userId'] }
    );
  }

  private async getUserName(args: CommandArgs): Promise<string> {
    if (args.options.userName) {
      return args.options.userName;
    }
    else {
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
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeUserfromSPGroup: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing User ${args.options.userName || args.options.email || args.options.userId} from Group: ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
        ? `GetById('${args.options.groupId}')`
        : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}`;

      if (args.options.userId) {
        requestUrl += `/users/removeById(${args.options.userId})`;
      }
      else {
        const userName: string = await this.getUserName(args);
        const loginName: string = `i:0#.f|membership|${userName}`;
        requestUrl += `/users/removeByLoginName(@LoginName)?@LoginName='${formatting.encodeQueryParameter(loginName)}'`;
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
    };

    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation bypassed by entering confirm option. Removing the user from SharePoint Group...');
      }
      await removeUserfromSPGroup();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove user User ${args.options.userName} from SharePoint group?`
      });

      if (result.continue) {
        await removeUserfromSPGroup();
      }
    }
  }
}

module.exports = new SpoGroupMemberRemoveCommand();