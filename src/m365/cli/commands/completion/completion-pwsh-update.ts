import { z } from 'zod';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape
});

class CliCompletionPwshUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_PWSH_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for PowerShell';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();
  }
}

export default new CliCompletionPwshUpdateCommand();