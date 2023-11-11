import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandError } from '../../Command.js';
import GlobalOptions from '../../GlobalOptions.js';
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


    try {
      if (args.options.identityId || args.options.identityName) {
        const identity = await auth.getIdentity(args.options.identityId, args.options.identityName);

        if (this.verbose) {
          await logger.logToStderr(`Logging out from identity ${identity.identityId}...`);
        }

        if (auth.service.identityId === identity.identityId) {
          auth.service.logout();
        }

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
}

export default new LogoutCommand();