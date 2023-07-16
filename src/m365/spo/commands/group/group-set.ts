import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  newName?: string;
  description?: string;
  allowMembersEditMembership?: boolean;
  onlyAllowMembersViewMembership?: boolean;
  allowRequestToJoinLeave?: boolean;
  autoAcceptRequestToJoinLeave?: boolean;
  requestToJoinLeaveEmailSetting?: string;
  ownerEmail?: string;
  ownerUserName?: string;
}

class SpoGroupSetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a group in the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        allowMembersEditMembership: args.options.allowMembersEditMembership,
        onlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        allowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        autoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        requestToJoinLeaveEmailSetting: typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined',
        ownerEmail: typeof args.options.ownerEmail !== 'undefined',
        ownerUserName: typeof args.options.ownerUserName !== 'undefined'
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
        option: '-n, --name [name]'
      },
      {
        option: '--newName [newName]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--allowMembersEditMembership [allowMembersEditMembership]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowRequestToJoinLeave [allowRequestToJoinLeave]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
      },
      {
        option: '--ownerEmail [ownerEmail]'
      },
      {
        option: '--ownerUserName [ownerUserName]'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowMembersEditMembership', 'onlyAllowMembersViewMembership', 'allowRequestToJoinLeave', 'autoAcceptRequestToJoinLeave');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id && isNaN(args.options.id)) {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.ownerEmail && args.options.ownerUserName) {
          return 'Specify either ownerEmail or ownerUserName but not both';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Setting properties for group ${args.options.id || args.options.name}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/sitegroups/${args.options.id ? `GetById(${args.options.id})` : `GetByName('${args.options.name}')`}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        Title: args.options.newName,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    try {
      await request.patch(requestOptions);
      if (args.options.ownerEmail || args.options.ownerUserName) {
        await this.setGroupOwner(args.options, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async setGroupOwner(options: Options, logger: Logger): Promise<void> {
    const ownerId = await this.getOwnerId(options, logger);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/sitegroups/${options.id ? `GetById(${options.id})` : `GetByName('${options.name}')`}/SetUserAsOwner(${ownerId})`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getOwnerId(options: Options, logger: Logger): Promise<number> {
    let userPrincipalName;
    if (options.ownerUserName) {
      userPrincipalName = options.ownerUserName;
    }
    else {
      userPrincipalName = await aadUser.getUpnByUserEmail(options.ownerEmail!, logger, this.verbose);
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/web/ensureUser('${userPrincipalName}')?$select=Id`,
      headers: {
        accept: 'application/json',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.post<{ Id: number }>(requestOptions);
    return response.Id;
  }
}

export default new SpoGroupSetCommand();