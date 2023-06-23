import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { aadUser } from '../../../../utils/aadUser';
import { validation } from '../../../../utils/validation';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  appName: string;
  userId?: string;
  userName?: string;
  roleForOldAppOwner?: string;
}

class PaAppOwnerSetCommand extends PowerAppsCommand {
  private static readonly roleForOldAppOwner = ['CanView', 'CanEdit'];
  public get name(): string {
    return commands.APP_OWNER_SET;
  }

  public get description(): string {
    return 'Sets a new owner for a Power Apps app';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        roleForOldAppOwner: typeof args.options.roleForOldAppOwner !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--appName <appName>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--roleForOldAppOwner [roleForOldAppOwner]',
        autocomplete: PaAppOwnerSetCommand.roleForOldAppOwner
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.appName)) {
          return `${args.options.appName} is not a valid GUID for appName`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for userId`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN for userName`;
        }

        if (args.options.roleForOldAppOwner && PaAppOwnerSetCommand.roleForOldAppOwner.indexOf(args.options.roleForOldAppOwner) < 0) {
          return `${args.options.roleForOldAppOwner} is not a valid roleForOldAppOwner. Allowed values are: ${PaAppOwnerSetCommand.roleForOldAppOwner.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Setting new owner ${args.options.userId || args.options.userName} for Power Apps app ${args.options.appName}...`);
    }
    try {
      const userId = await this.getUserId(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.appName}/modifyAppOwner?api-version=2022-11-01`,
        headers: {
          accept: 'application/json',
          'Content-Type': 'application/json'
        },
        responseType: 'json',
        data: {
          roleForOldAppOwner: args.options.roleForOldAppOwner,
          newAppOwner: userId
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(options: Options): Promise<string> {
    if (options.userId) {
      return options.userId;
    }

    return aadUser.getUserIdByUpn(options.userName!);
  }
}

module.exports = new PaAppOwnerSetCommand();