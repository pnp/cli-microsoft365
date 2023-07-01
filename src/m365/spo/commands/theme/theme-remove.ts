import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  force?: boolean;
}

class SpoThemeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_REMOVE;
  }

  public get description(): string {
    return 'Removes existing theme';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeTheme(logger, args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the theme`
      });

      if (result.continue) {
        await this.removeTheme(logger, args.options);
      }
    }
  }

  private async removeTheme(logger: Logger, options: GlobalOptions): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Removing theme from tenant...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/thememanager/DeleteTenantTheme`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          name: options.name
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoThemeRemoveCommand();