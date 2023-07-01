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
  scope?: string
}

class SpoApplicationCustomizerListCommand extends SpoCommand {
  private static readonly scopes: string[] = ['All', 'Site', 'Web'];

  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_LIST;
  }

  public get description(): string {
    return 'Get a list of application customizers that are added to a site.';
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

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpoApplicationCustomizerListCommand.scopes
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope && SpoApplicationCustomizerListCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpoApplicationCustomizerListCommand.scopes.join(', ')}`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving application customizers...`);
    }

    const applicationCustomizers = await spo.getCustomActions(args.options.webUrl, args.options.scope, `Location eq 'ClientSideExtension.ApplicationCustomizer'`);
    await logger.log(applicationCustomizers);
  }
}

export default new SpoApplicationCustomizerListCommand();