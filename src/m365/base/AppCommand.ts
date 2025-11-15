import fs from 'fs';
import { z } from 'zod';
import { cli } from '../../cli/cli.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandError, globalOptionsZod } from '../../Command.js';
import { formatting } from '../../utils/formatting.js';
import { M365RcJson, M365RcJsonApp } from './M365RcJson.js';

export const appCommandOptions = z.object({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional()
});
type Options = z.infer<typeof appCommandOptions>;

export interface AppCommandArgs {
  options: Options;
}

export default abstract class AppCommand extends Command {
  protected m365rcJson: M365RcJson | undefined;
  protected appId: string | undefined;

  protected get resource(): string {
    return 'https://graph.microsoft.com';
  }

  public get schema(): z.ZodType | undefined {
    return appCommandOptions;
  }

  public async action(logger: Logger, args: AppCommandArgs): Promise<void> {
    const m365rcJsonPath: string = '.m365rc.json';

    if (!fs.existsSync(m365rcJsonPath)) {
      throw new CommandError(`Could not find file: ${m365rcJsonPath}`);
    }

    try {
      const m365rcJsonContents: string = fs.readFileSync(m365rcJsonPath, 'utf8');
      if (!m365rcJsonContents) {
        throw new CommandError(`File ${m365rcJsonPath} is empty`);
      }

      this.m365rcJson = JSON.parse(m365rcJsonContents) as M365RcJson;
    }
    catch (err) {
      if (err instanceof CommandError) {
        throw err;
      }
      throw new CommandError(`Could not parse file: ${m365rcJsonPath}`);
    }

    if (!this.m365rcJson.apps ||
      this.m365rcJson.apps.length === 0) {
      throw new CommandError(`No Entra apps found in ${m365rcJsonPath}`);
    }

    if (args.options.appId) {
      if (!this.m365rcJson.apps.some(app => app.appId === args.options.appId)) {
        throw new CommandError(`App ${args.options.appId} not found in ${m365rcJsonPath}`);
      }

      this.appId = args.options.appId;
      return super.action(logger, args);
    }

    if (this.m365rcJson.apps.length === 1) {
      this.appId = this.m365rcJson.apps[0].appId;
      return super.action(logger, args);
    }

    if (this.m365rcJson.apps.length > 1) {
      this.m365rcJson.apps.forEach((app, index) => {
        (app as any).appIdIndex = index;
      });
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('appId', this.m365rcJson.apps);
      const result = await cli.handleMultipleResultsFound<M365RcJsonApp>(`Multiple Entra apps found in ${m365rcJsonPath}.`, resultAsKeyValuePair);
      this.appId = result.appId;
      await super.action(logger, args);
    }
  }
}