import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  scope?: string;
}

class SpoCommandsetListCommand extends SpoCommand {
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
        autocomplete: SpoCommandsetListCommand.scopes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope && SpoCommandsetListCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Valid scopes are ${SpoCommandsetListCommand.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Attempt to get commandsets...`);
      }

      const commandsets = await spo.getCustomActions(args.options.webUrl, args.options.scope, `startswith(Location,'ClientSideExtension.ListViewCommandSet')`);

      logger.log(commandsets);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoCommandsetListCommand();