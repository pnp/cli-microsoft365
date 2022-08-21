import * as fs from 'fs';
import { Cli, Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
import GlobalOptions from '../../GlobalOptions';
import { validation } from '../../utils';
import { M365RcJson, M365RcJsonApp } from './M365RcJson';

export interface AppCommandArgs {
  options: AppCommandOptions;
}

export interface AppCommandOptions extends GlobalOptions {
  appId?: string;
}

export default abstract class AppCommand extends Command {
  protected m365rcJson: M365RcJson | undefined;
  protected appId: string | undefined;

  protected get resource(): string {
    return 'https://graph.microsoft.com';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        return true;
      },
    );
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
    catch (e) {
      throw new CommandError(`Could not parse file: ${m365rcJsonPath}`);
    }

    if (!this.m365rcJson.apps ||
      this.m365rcJson.apps.length === 0) {
      throw new CommandError(`No Azure AD apps found in ${m365rcJsonPath}`);
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
      const result = await Cli.prompt<{ appIdIndex: number }>({
        message: `Multiple Azure AD apps found in ${m365rcJsonPath}. Which app would you like to use?`,
        type: 'list',
        choices: this.m365rcJson.apps.map((app, i) => {
          return {
            name: `${app.name} (${app.appId})`,
            value: i
          };
        }),
        default: 0,
        name: 'appIdIndex'
      });

      this.appId = ((this.m365rcJson as M365RcJson).apps as M365RcJsonApp[])[result.appIdIndex].appId;
      await super.action(logger, args);
    }
  }
}