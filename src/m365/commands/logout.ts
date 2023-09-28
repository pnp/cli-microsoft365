import auth, { Identity } from '../../Auth.js';
import { Cli } from '../../cli/Cli.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandError } from '../../Command.js';
import GlobalOptions from '../../GlobalOptions.js';
import { formatting } from '../../utils/formatting.js';
import { validation } from '../../utils/validation.js';
import commands from './commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  identityId?: string;
  identityName?: string;
}

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initValidators();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        identityId: typeof args.options.identityId !== 'undefined',
        identityName: typeof args.options.identityName !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.identityId && !validation.isValidGuid(args.options.identityId as string)) {
          return `${args.options.identityId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --identityId [identityId]'
      },
      {
        option: '-n, --identityName [identityName]'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['identityId', 'identityName'], runsWhen: (args) => args.options.identityId || args.options.identityName });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Logging out from Microsoft 365...');
    }

    const identity = await this.getIdentityToLogout(args.options);

    try {
      if (identity) {
        if (this.verbose) {
          await logger.logToStderr(`Logging out from identity ${identity.identityId}...`);
        }

        auth.service.logout(identity.identityId);

        await auth.clearConnectionInfo(logger, this.debug, identity.identityId);
      }
      else {
        auth.service.logout();

        await auth.clearConnectionInfo(logger, this.debug);
      }
    }
    catch (error: any) {
      if (this.debug) {
        await logger.logToStderr(new CommandError(error));
      }
    }
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    this.initAction(args, logger);
    await this.commandAction(logger, args);
  }

  private async getIdentityToLogout(options: Options): Promise<Identity | undefined> {
    try {
      if (!options.identityId && !options.identityName) {
        return;
      }

      const identities = auth.service.availableIdentities!.filter(i => i.identityName === options.identityName || i.identityId === options.identityId);

      if (identities.length === 0) {
        throw new Error(`The identity '${options.identityId || options.identityName}' cannot be found.`);
      }

      if (identities.length > 1) {
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('identityId', identities);
        const result = await Cli.handleMultipleResultsFound<Identity>(`Multiple identities with '${options.identityName}' found.`, resultAsKeyValuePair);
        return result;
      }

      return identities[0];
    }
    catch (error: any) {
      throw new CommandError(error.message);
    }
  }
}

export default new LogoutCommand();