import { Application } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from "../../../../utils/odata.js";
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import GlobalOptions from '../../../../GlobalOptions.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  properties?: string;
}

class EntraAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of Entra app registrations';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  public alias(): string[] | undefined {
    return [aadCommands.APP_LIST, commands.APPREGISTRATION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'id', 'displayName', "signInAudience"];
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-p, --properties [properties]' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const queryParameters: string[] = [];

    if (args.options.properties) {
      const allProperties = args.options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    try {
      const results = await odata.getAllItems<Application>(`${this.resource}/v1.0/applications${queryString}`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAppListCommand();