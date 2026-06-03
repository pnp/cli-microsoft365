import { z } from 'zod';
import { globalOptionsZod } from '../../../Command.js';
import ContextCommand from '../../base/ContextCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({ ...globalOptionsZod.shape });

class ContextInitCommand extends ContextCommand {
  public get name(): string {
    return commands.INIT;
  }

  public get description(): string {
    return 'Initiates CLI for Microsoft 365 context in the current working folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(): Promise<void> {
    await this.saveContextInfo({});
  }
}

export default new ContextInitCommand();