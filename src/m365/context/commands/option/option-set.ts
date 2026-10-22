import fs from 'fs';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import ContextCommand from '../../../base/ContextCommand.js';
import { M365RcJson } from '../../../base/M365RcJson.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  value: z.string().alias('v')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ContextOptionSetCommand extends ContextCommand {
  public get name(): string {
    return commands.OPTION_SET;
  }

  public get description(): string {
    return 'Allows to add a new name for the option and value to the local context file.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      await logger.logToStderr(`Saving ${args.options.name} with value ${args.options.value} to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Reading existing ${filePath}...`);
        }

        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        throw new CommandError(`Error reading ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
      }
    }

    if (m365rc.context) {
      m365rc.context[args.options.name] = args.options.value;
      try {
        if (this.verbose) {
          await logger.logToStderr(`Creating option ${args.options.name} with value ${args.options.value} in existing context...`);
        }
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
      }
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Context doesn't exist. Initializing the context and creating option ${args.options.name} with value ${args.options.value}...`);
      }

      this.saveContextInfo({ [args.options.name]: args.options.value });
    }
  }
}

export default new ContextOptionSetCommand();