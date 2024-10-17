
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

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
    return ['name', 'deploymentpipelineid', 'ownerid', 'statuscode'];
  }
  private async getEnvironmentDetails(environmentName: string, asAdmin: boolean): Promise<any> {
    let url: string = `${this.resource}/providers/Microsoft.BusinessAppPlatform`;
    if (asAdmin) {
      url += '/scopes/admin';
    }

    const envName = formatting.encodeQueryParameter(environmentName);
    url += `/environments/${envName}?api-version=2020-10-01`;

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<any>(requestOptions);
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
    return pipelines.value.map((p: any) => {

      return {
        name: p.name,
        deploymentpipelineid: p.deploymentpipelineid,
        ownerid: p['_ownerid_value'],
        statuscode: p.statuscode
      };
    });
  }
  public async commandAction(logger: Logger, args: any): Promise<void> {

    try {
      const environmentDetails = await this.getEnvironmentDetails(args.options.environmentName, args.options.asAdmin);
      const instanceUrl = environmentDetails.properties.linkedEnvironmentMetadata.instanceApiUrl;

      const pipelines = await this.listPipelines(instanceUrl);
      await logger.log(pipelines);
    }
    catch (ex: any) {
      this.handleRejectedODataJsonPromise(ex);
    }


  }

}


export default new PpPipelineListCommand();