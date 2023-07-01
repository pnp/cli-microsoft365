import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  scope?: string;
}

class SpoCommandSetListCommand extends SpoCommand {
  private static readonly scopes: string[] = ['All', 'Site', 'Web'];
  public get name(): string {
    return commands.COMMANDSET_LIST;
  }

  public get description(): string {
    return 'Get a list of ListView Command Sets that are added to a site.';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        scope: args.options.scope || 'All'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpoCommandSetListCommand.scopes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope && SpoCommandSetListCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Valid scopes are ${SpoCommandSetListCommand.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Attempt to get commandsets...`);
      }

      const commandsets = await spo.getCustomActions(args.options.webUrl, args.options.scope, `startswith(Location,'ClientSideExtension.ListViewCommandSet')`);

      await logger.log(commandsets);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoCommandSetListCommand();