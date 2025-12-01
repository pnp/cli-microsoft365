import fs from 'fs';
import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../Command.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import { M365RcJson } from '../../base/M365RcJson.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ContextRemoveCommand extends AnonymousCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the CLI for Microsoft 365 context in the current working folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      this.removeContext();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the context?` });

      if (result) {
        this.removeContext();
      }
    }
  }

  private removeContext(): void {
    const filePath: string = '.m365rc.json';

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        throw new CommandError(`Error reading ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
      }
    }

    if (!m365rc.context) {
      return;
    }

    const keys = Object.keys(m365rc);
    if (keys.length === 1 && keys.indexOf('context') > -1) {
      try {
        fs.unlinkSync(filePath);
      }
      catch (e) {
        throw new CommandError(`Error removing ${filePath}: ${e}. Please remove ${filePath} manually.`);
      }
    }
    else {
      try {
        delete m365rc.context;
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
      }
    }
  }
}

export default new ContextRemoveCommand();