// import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
// import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { TeamsApp } from '../../TeamsApp';
// import { formatting } from '../../../../utils/formatting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  distributionMethod?: string;
}

class TeamsAppListCommand extends GraphCommand {
  private static allowedDistributionMethods: string[] = ['store', 'organization', 'sideloaded'];

  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the Microsoft Teams app catalog';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
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
        distributionMethod: typeof args.options.distributionMethod !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--distributionMethod',
        autocomplete: TeamsAppListCommand.allowedDistributionMethods
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.distributionMethod &&
          TeamsAppListCommand.allowedDistributionMethods.indexOf(args.options.distributionMethod) < 0) {
          return `'${args.options.distributionMethod}' is not a valid distribution method. Allowed distribution methods are ${TeamsAppListCommand.allowedDistributionMethods.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestUrl: string = `${this.resource}/v1.0/appCatalogs/teamsApps${args.options.distributionMethod ? `?$filter=distributionMethod eq '${args.options.distributionMethod}'` : ''}`;

      const items = await odata.getAllItems<TeamsApp>(requestUrl);

      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsAppListCommand();