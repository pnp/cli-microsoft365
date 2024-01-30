import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
        id: (!(!args.options.id)).toString(),
        loginName: (!(!args.options.loginName)).toString(),
        email: (!(!args.options.email)).toString(),
        userName: (!(!args.options.userName)).toString(),
        entraGroupId: (!(!args.options.entraGroupId)).toString(),
        entraGroupName: (!(!args.options.entraGroupName)).toString(),
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

        if (args.options.id &&
          typeof args.options.id !== 'number') {
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
      else if (options.email || options.userName) {
        const user = await this.getUser(options.webUrl, options);

        if (user?.Title) {
          if (this.verbose) {
            await logger.logToStderr(`Removing user ${user.Title} ...`);
          }
          requestUrl += `removebyid(${user.Id})`;
        }
        else {
          throw new Error(`User not found: ${options.userName}`);
        }
      }
      else if (options.entraGroupId || options.entraGroupName) {
        const entraGroup = await this.getEntraGroup(options.webUrl, options);
        if (this.verbose) {
          await logger.logToStderr(`Removing entra group ${entraGroup.displayName} ...`);
        }
        //for entra groups, M365 groups have an associated email and security groups don't
        if (entraGroup.mail) {
          //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
          requestUrl += `removeByLoginName('c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}')`;
        }
        else {
          //security group is prefixed with c:0t.c|tenant
          requestUrl += `removeByLoginName('c:0t.c|tenant|${entraGroup.id}')`;
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

  private async getUser(webUrl: string, options: GlobalOptions): Promise<any> {
    let requestUrl: string = '';
    if (options.email) {
      requestUrl = `${options.webUrl}/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter(options.email)}')`;
    }
    else if (options.userName) {
      requestUrl = `${options.webUrl}/_api/web/siteusers?$filter=UserPrincipalName eq ('${formatting.encodeQueryParameter(options.userName)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const userInstance = await request.get(requestOptions);
    if (options.email) {
      return userInstance;
    }
    else if (options.userName) {
      return (userInstance as { value: any[] }).value[0];
    }
  }

  private async getEntraGroup(webUrl: string, options: GlobalOptions): Promise<any> {
    let group: Group;
    if (options.entraGroupId) {
      group = await aadGroup.getGroupById(options.entraGroupId);
      return group;
    }
    else if (options.entraGroupName) {
      group = await aadGroup.getGroupByDisplayName(options.entraGroupName);
      return group;
    }

  }
}

export default new SpoUserRemoveCommand();