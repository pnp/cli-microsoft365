import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CommandError } from '../../../../Command.js';

interface CommandArgs {
  options: Options;
}
interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  loginName?: string;
  email?: string;
  userName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
  force: boolean;
}

class SpoUserRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes user from specific web';
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
        id: typeof args.options.id !== 'undefined',
        loginName: typeof args.options.loginName !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--loginName [loginName]'
      },
      {
        option: '--email [email]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
      },
      {
        option: '-f, --force'
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

        if (args.options.id && isNaN(parseInt(args.options.id))) {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `${args.options.entraId} is not a valid GUID.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email.`;
        }
        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['id', 'loginName', 'email', 'userName', 'entraGroupId', 'entraGroupName']
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeUser(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified user from the site ${args.options.webUrl}?` });

      if (result) {
        await this.removeUser(logger, args.options);
      }
    }
  }

  private async removeUser(logger: Logger, options: GlobalOptions): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing user from  subsite ${options.webUrl} ...`);
    }
    try {
      let requestUrl: string = `${encodeURI(options.webUrl)}/_api/web/siteusers/`;
      if (options.id) {
        requestUrl += `removebyid(${options.id})`;
      }
      else if (options.loginName) {
        requestUrl += `removeByLoginName('${formatting.encodeQueryParameter(options.loginName as string)}')`;
      }
      else if (options.email) {
        const user = await spo.getUserByEmail(options.webUrl, options.email, logger, this.verbose);
        requestUrl += `removebyid(${user.Id})`;
      }
      else if (options.userName) {
        const user = await this.getUser(options);

        if (!user) {
          throw new CommandError(`User not found: ${options.userName}`);
        }

        if (this.verbose) {
          await logger.logToStderr(`Removing user ${user.Title} ...`);
        }
        requestUrl += `removebyid(${user.Id})`;
      }
      else if (options.entraGroupId || options.entraGroupName) {
        const entraGroup = await this.getEntraGroup(options.webUrl, options);
        if (this.verbose) {
          await logger.logToStderr(`Removing entra group ${entraGroup?.displayName} ...`);
        }
        //for entra groups, M365 groups have an associated email and security groups don't
        if (entraGroup?.mail) {
          //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
          requestUrl += `removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}')`;
        }
        else {
          //security group is prefixed with c:0t.c|tenant
          requestUrl += `removeByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
        }
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUser(options: GlobalOptions): Promise<any> {
    const requestUrl: string = `${options.webUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(options.userName)}')`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get(requestOptions);
    return (userInstance as {
      value: spoUser[];
    }).value[0];
  }

  private async getEntraGroup(webUrl: string, options: GlobalOptions): Promise<Group | undefined> {
    let group: Group | undefined;
    if (options.entraGroupId) {
      group = await entraGroup.getGroupById(options.entraGroupId);
    }
    else if (options.entraGroupName) {
      group = await entraGroup.getGroupByDisplayName(options.entraGroupName);
    }
    return group;
  }
}

export default new SpoUserRemoveCommand();

interface spoUser {
  Id: number;
  IsHiddenInUI: boolean;
  Title: string;
  PrincipalType: number;
  Email: string;
  Expiration: string;
  IsEmailAuthenticationGuestUser: boolean;
  IsShareByEmailGuestUser: boolean;
  IsSiteAdmin: boolean;
  UserId: {
    NameId: string;
    NameIdIssuer: string;
    urn: string;
  };
  UserPrincipalName: string;
};
// Add your code or comments here
