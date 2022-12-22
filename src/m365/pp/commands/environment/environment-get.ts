import { Logger } from '../../../../cli/Logger';
import { AxiosRequestConfig } from 'axios';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import { Environment } from '../Environment';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  asAdmin?: boolean;
}

class PpEnvironmentGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Power Platform environment';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'id'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments`;
    if (args.options.asAdmin) {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${url}?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const res: { value: Environment[] } = await request.get<{ value: Environment[] }>(requestOptions);
    const environmentItem: Environment | undefined = res.value.filter((env: Environment) => {
      return args.options.name ? env.name === args.options.name : env.properties.isDefault === true;
    })[0];

    if (!environmentItem) {
      throw `The specified Power Platform environment does not exist`;
    }

    logger.log(environmentItem);
  }
}

module.exports = new PpEnvironmentGetCommand();