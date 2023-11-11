import { Logger } from '../../../cli/Logger.js';
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError } from '../../../Command.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { validation } from '../../../utils/validation.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class IdentitySetCommand extends Command {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return "Switches to another identity, when signed into multiple identities";
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const identity = await auth.getIdentity(args.options.id, args.options.name);

    if (this.verbose) {
      await logger.logToStderr(`Switching to identity '${identity.identityName}'...`);
    }

    await auth.switchToConnection(identity);

    await logger.log(auth.getIdentityDetails(auth.service, this.debug));
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

export default new IdentitySetCommand();