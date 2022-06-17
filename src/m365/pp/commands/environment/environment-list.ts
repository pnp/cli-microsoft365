import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  asAdmin?: boolean;
}

class PpEnvironmentListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform environments';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-a, --asAdmin [teamId]'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of Microsoft Power Platform environments...`);
    }
    let url: string = '';
    if (args.options.asAdmin) {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`;
    }
    else {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments`;
    }

    const requestOptions: any = {
      url: `${url}?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ name: string, displayName: string; properties: { displayName: string } }] }>(requestOptions)
      .then((res: { value: [{ name: string, displayName: string; properties: { displayName: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          res.value.forEach(e => {
            e.displayName = e.properties.displayName;
          });

          logger.log(res.value);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new PpEnvironmentListCommand();
