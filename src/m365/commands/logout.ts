import { z } from 'zod';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandError, globalOptionsZod } from '../../Command.js';
import commands from './commands.js';

export const options = globalOptionsZod.strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Logging out from Microsoft 365...');
    }

    const deactivate: () => void = (): void => auth.connection.deactivate();

    try {
      await auth.clearConnectionInfo();
    }
    catch (error: any) {
      if (this.debug) {
        await logger.logToStderr(new CommandError(error));
      }
    }
    finally {
      deactivate();
    }
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger);
  }
}

export default new LogoutCommand();