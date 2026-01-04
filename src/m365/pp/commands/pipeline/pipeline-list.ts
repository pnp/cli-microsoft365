
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface Options extends GlobalOptions {
  environmentName: string;
  asAdmin?: boolean;
}

interface CommandArgs {
  options: Options;
}

class PpPipelineListCommand extends PowerPlatformCommand {

  constructor() {
    super();
    this.#initTelemetry();
    this.#initOptions();
  }

  public get name(): string {
    return commands.PIPELINE_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform pipelines in the specified Power Platform environment.';
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
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'deploymentpipelineid', '_ownerid_value', 'statuscode'];
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const pipelines = await this.listPipelines(dynamicsApiUrl);

      await logger.log(pipelines);
    }
    catch (ex: any) {
      this.handleRejectedODataJsonPromise(ex);
    }

  }

  private async listPipelines(instanceUrl: string): Promise<any> {

    const pipelineListRequestOptions: CliRequestOptions = {
      url: `${instanceUrl}/api/data/v9.0/deploymentpipelines`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const pipelines = await request.get<any>(pipelineListRequestOptions);

    return pipelines.value;
  }

}

export default new PpPipelineListCommand();