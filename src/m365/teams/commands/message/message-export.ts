import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import * as fs from 'fs';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  teamId?: string;
  fromDateTime?: string;
  toDateTime?: string;
  licenseModel?: string;
  withAttachments?: boolean;
  folderPath: string;
}

class TeamsMessageExportCommand extends GraphCommand {
  private readonly allowedLicenseModels: string[] = ['A', 'B'];

  public get name(): string {
    return commands.MESSAGE_EXPORT;
  }

  public get description(): string {
    return 'Export Microsoft Teams chat messages for a given user, or a team.';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--fromDateTime [fromDateTime]'
      },
      {
        option: '--toDateTime [toDateTime]'
      },
      {
        option: '--licenseModel [licenseModel]',
        autocomplete: this.allowedLicenseModels
      },
      {
        option: '--withAttachments'
      },
      {
        option: '--folderPath <folderPath>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userId} is not a valid userPrincipalName`;
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.fromDateTime && !validation.isValidISODateTime(args.options.fromDateTime)) {
          return `${args.options.fromDateTime} is not a valid ISO DateTime`;
        }

        if (args.options.toDateTime && !validation.isValidISODateTime(args.options.toDateTime)) {
          return `${args.options.toDateTime} is not a valid ISO DateTime`;
        }

        if (args.options.licenseModel && !this.allowedLicenseModels.some(value => value === args.options.licenseModel)) {
          return `${args.options.licenseModel} is not a valid license model. Allowed values are ${this.allowedLicenseModels.join(',')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'teamId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!fs.existsSync(args.options.folderPath)) {
      throw `Path ${args.options.folderPath} does not exist.`;
    }

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (args.options.userId || args.options.userName) {
      requestOptions.url = `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/chats`;
    }
    else {
      requestOptions.url = `${this.resource}/v1.0/teams/${args.options.teamId}/channels`;
    }
    requestOptions.url += '/getAllMessages';

    try {
      const res: any = await request.get(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMessageExportCommand();