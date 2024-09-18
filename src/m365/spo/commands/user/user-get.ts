import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface SpoUser {
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
}

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  email?: string;
  loginName?: string;
  userName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
}

class SpoUserGetCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets a site user within specific web';
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
        id: typeof args.options.id !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        loginName: typeof args.options.loginName !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined'
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
        option: '--email [email]'
      },
      {
        option: '--loginName [loginName]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'id', 'email', 'loginName', 'userName', 'entraGroupId', 'entraGroupName');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id &&
          typeof args.options.id !== 'number') {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `${args.options.entraGroupId} is not a valid GUID.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email.`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['id', 'email', 'loginName', 'userName', 'entraGroupId', 'entraGroupName'],
      runsWhen: (args) => args.options.id || args.options.email || args.options.loginName || args.options.userName || args.options.entraGroupId || args.options.entraGroupName
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information for user in site '${args.options.webUrl}'...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    if (args.options.id) {
      requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(args.options.id.toString())}')`;
    }
    else if (args.options.email) {
      requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(args.options.email)}')`;
    }
    else if (args.options.loginName) {
      requestUrl += `siteusers/GetByLoginName('${formatting.encodeQueryParameter(args.options.loginName)}')`;
    }
    else if (args.options.userName) {
      const user = await this.getUser(requestUrl, args.options.userName);
      requestUrl += `siteusers/GetById('${formatting.encodeQueryParameter(user.Id.toString())}')`;
    }
    else if (args.options.entraGroupId || args.options.entraGroupName) {
      const entraGroup = await this.getEntraGroup(args.options.entraGroupId!, args.options.entraGroupName!);

      // For entra groups, M365 groups have an associated email and security groups don't
      if (entraGroup?.mail) {
        requestUrl += `siteusers/GetByEmail('${formatting.encodeQueryParameter(entraGroup.mail)}')`;
      }
      else {
        requestUrl += `siteusers/GetByLoginName('c:0t.c|tenant|${entraGroup?.id}')`;
      }
    }
    else {
      requestUrl += `currentuser`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'GET',
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const userInstance = await request.get(requestOptions);
      await logger.log(userInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUser(baseUrl: string, userName: string): Promise<SpoUser> {
    const requestUrl: string = `${baseUrl}siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(userName)}')`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get(requestOptions);
    const userInstanceValue = (userInstance as {
      value: SpoUser[];
    }).value[0];

    if (!userInstanceValue) {
      throw `User not found: ${userName}`;
    }

    return userInstanceValue;
  }

  private async getEntraGroup(entraGroupId: string, entraGroupName: string): Promise<Group> {
    if (entraGroupId) {
      return entraGroup.getGroupById(entraGroupId);
    }

    return entraGroup.getGroupByDisplayName(entraGroupName);
  }
}

export default new SpoUserGetCommand();