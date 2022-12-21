import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin?: boolean;
}

class PpAibuildermodelListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.AIBUILDERMODEL_LIST;
  }

  public get description(): string {
    return 'List available AI builder models in the specified Power Platform environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['msdyn_name', 'msdyn_aimodelid', 'createdon', 'modifiedon'];
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
        option: '-e, --environment <environment>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving available AI builder models`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const aimodels = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.0/msdyn_aimodels?$filter=iscustomizable/Value eq true`);
      logger.log(aimodels);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpAibuildermodelListCommand();