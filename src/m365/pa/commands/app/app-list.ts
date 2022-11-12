import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment?: string;
  asAdmin: boolean;
}

class PaAppListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists all Power Apps apps';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
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
        asAdmin: args.options.asAdmin === true,
        environment: typeof args.options.environment !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment [environment]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.asAdmin && !args.options.environment) {
          return 'When specifying the asAdmin option the environment option is required as well';
        }

        if (args.options.environment && !args.options.asAdmin) {
          return 'When specifying the environment option the asAdmin option is required as well';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environment ? '/environments/' + formatting.encodeQueryParameter(args.options.environment) : ''}/apps?api-version=2017-08-01`;

    try {
      const apps = await odata.getAllItems<{ name: string; displayName: string; properties: { displayName: string } }>(url);

      if (apps.length > 0) {
        apps.forEach(a => {
          a.displayName = a.properties.displayName;
        });

        logger.log(apps);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No apps found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaAppListCommand();