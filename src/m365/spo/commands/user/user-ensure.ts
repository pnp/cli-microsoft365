import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  entraId?: string;
  userName?: string;
  loginName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
}

class SpoUserEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_ENSURE;
  }

  public get description(): string {
    return 'Ensures that a user is available on a specific site';
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
        entraId: typeof args.options.entraId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        loginName: typeof args.options.loginName !== 'undefined',
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
        option: '--entraId [entraId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--loginName [loginName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
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

        if (args.options.entraId && !validation.isValidGuid(args.options.entraId)) {
          return `${args.options.entraId} is not a valid GUID.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `${args.options.entraGroupId} is not a valid GUID for option 'entraGroupId'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['entraId', 'userName', 'loginName', 'entraGroupId', 'entraGroupName'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'entraId', 'userName', 'loginName', 'entraGroupId', 'entraGroupName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Ensuring user ${args.options.entraId || args.options.userName || args.options.loginName || args.options.entraGroupId || args.options.entraGroupName} at site ${args.options.webUrl}`);
    }

    try {
      const requestBody = {
        logonName: await this.getUpn(args.options)
      };

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/ensureuser`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUpn(options: Options): Promise<string> {
    if (options.userName) {
      return options.userName;
    }

    if (options.entraId) {
      return entraUser.getUpnByUserId(options.entraId);
    }

    if (options.loginName) {
      return options.loginName;
    }

    let upn: string = '';
    if (options.entraGroupId || options.entraGroupName) {
      const entraGroup = await this.getEntraGroup(options.entraGroupId, options.entraGroupName);
      upn = entraGroup.mailEnabled ? `c:0o.c|federateddirectoryclaimprovider|${entraGroup.id}` : `c:0t.c|tenant|${entraGroup.id}`;
    }

    return upn;
  }

  private async getEntraGroup(entraGroupId?: string, entraGroupName?: string): Promise<Group> {
    if (entraGroupId) {
      return entraGroup.getGroupById(entraGroupId);
    }

    return entraGroup.getGroupByDisplayName(entraGroupName!);
  }
}

export default new SpoUserEnsureCommand();
