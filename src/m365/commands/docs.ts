import { z } from 'zod';
import { cli } from '../../cli/cli.js';
import { Logger } from '../../cli/Logger.js';
import { globalOptionsZod } from '../../Command.js';
import { settingsNames } from '../../settingsNames.js';
import { app } from '../../utils/app.js';
import { browserUtil } from '../../utils/browserUtil.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';

export const options = globalOptionsZod.strict();

class DocsCommand extends AnonymousCommand {
  public get name(): string {
    return commands.DOCS;
  }

  public get description(): string {
    return 'Returns the CLI for Microsoft 365 docs webpage URL';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(app.packageJson().homepage);

    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      return;
    }

    await browserUtil.open(app.packageJson().homepage);
  }
}

export default new DocsCommand();