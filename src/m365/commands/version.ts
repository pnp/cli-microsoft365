import commands from '../commands.js';
import { Logger } from '../../../cli/Logger.js';
import Command, { CommandArgs, CommandOutput } from '../../Command.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../Command.js';
import app from '../../../utils/app.js';

/**
 * Zod schema for the version command options.
 * No additional options, just global options.
 */
export const options = globalOptionsZod.strict();
type Options = z.infer<typeof options>;

interface VersionCommandArgs extends CommandArgs {
  options: Options;
}

class VersionCommand extends Command {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI version';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: VersionCommandArgs): Promise<CommandOutput> {
    await logger.log(`v${app.packageJson().version}`);
    return;
  }
}

export default new VersionCommand();