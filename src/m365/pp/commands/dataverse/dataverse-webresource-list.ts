import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { odata } from '../../../../utils/odata.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { SolutionComponent } from './SolutionComponent.js';
import { COMPONENT_STATE_LABELS, WEBRESOURCE_TYPE_LABELS, Webresource } from './WebResource.js';
import { Solution } from '../solution/Solution.js';
import { validation } from '../../../../utils/validation.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  solutionId: string;
  excludeContent: boolean;
}

class PpDataverseWebResourceListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.DATAVERSE_WEBRESOURCE_LIST;
  }

  public get description(): string {
    return 'Lists web resources in the specified Power Platform solution';
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'webresourceType', 'webresourceTypeLabel', 'isManaged', 'isManagedLabel', 'canBeDeleted'];
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
        excludeContent: !!args.options.excludeContent
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.environmentName && !validation.isValidGuid(args.options.environmentName)) {
          return `The value provided as environmentName '${args.options.environmentName}' is not a valid GUID`;
        }

        if (args.options.solutionId && !validation.isValidGuid(args.options.solutionId)) {
          return `The value provided as solutionId '${args.options.solutionId}' is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-s, --solutionId <solutionId>'
      },
      {
        option: '--excludeContent'
      },
    );
  }

  private WEBRESOURCE_PROPS_ALL = [
    'webresourceid',
    'name',
    'canbedeleted',
    'componentstate',
    'content',
    'content_binary',
    'contentfileref',
    'contentjson',
    'contentjsonfileref',
    'createdon',
    'dependencyxml',
    'description',
    'displayname',
    'introducedversion',
    'isavailableformobileoffline',
    'iscustomizable',
    'isenabledformobileclient',
    'ishidden',
    'ismanaged',
    'languagecode',
    'modifiedon',
    'overwritetime',
    'silverlightversion',
    'solutionid',
    'versionnumber',
    'webresourceidunique',
    'webresourcetype'
  ];

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.logIfVerbose(logger, `Retrieving list of web resources for solution '${args.options.solutionId}' in environment '${args.options.environmentName}'...`);

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName);

      this.logIfVerbose(logger, `Found Dynamics API URL: ${dynamicsApiUrl}`);

      if (await this.checkIfSolutionExists(args.options.solutionId, dynamicsApiUrl, logger) === false) {
        throw new Error(`Solution with ID '${args.options.solutionId}' not found in environment '${args.options.environmentName}'`);
      }

      const solutionComponentsRequestUrl = `${dynamicsApiUrl}/api/data/v9.2/solutioncomponents?$filter=_solutionid_value eq ${args.options.solutionId} and componenttype eq 61&$select=objectid`;

      this.logIfVerbose(logger, `Retrieving web resources solution components`);
      const solutionComponentResponse = await odata.getAllItems<SolutionComponent>(solutionComponentsRequestUrl);

      this.logIfVerbose(logger, `Found ${solutionComponentResponse.length} web resource solution component in solution '${args.options.solutionId}' in environment '${args.options.environmentName}'`);
      const props = this.getWebresourceProperties(args);

      this.logIfVerbose(logger, `Using props: ${props.join(',')}`);
      const webResourcesRequestUrl = `${dynamicsApiUrl}/api/data/v9.2/webresourceset?$select=${props.join(',')}&$filter=${solutionComponentResponse.map((c: SolutionComponent) => `webresourceid eq ${c.objectid}`).join(' or ')}`;

      this.logIfVerbose(logger, `Retrieving web resources from URL: ${webResourcesRequestUrl}`);
      const webResourcesResponse = await odata.getAllItems<Webresource>(webResourcesRequestUrl);

      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        const outputWithLabels = webResourcesResponse.map(webresource => {
          return {
            ...webresource,
            webresourceTypeLabel: WEBRESOURCE_TYPE_LABELS[webresource.webresourcetype - 1],
            componentStateLabel: COMPONENT_STATE_LABELS[webresource.componentstate],
            isManagedLabel: webresource.ismanaged ? 'Managed' : 'Unmanaged',
            isEnabledForMobileClientLabel: webresource.isenabledformobileclient ? 'Yes' : 'No',
            isAvailableForMobileOfflineLabel: webresource.isavailableformobileoffline ? 'Yes' : 'No'
          };
        });
        await logger.log(outputWithLabels);
      }
      else {
        await logger.log(webResourcesResponse.map(i => {
          return {
            displayName: i.displayname,
            webresourceType: i.webresourcetype,
            webresourceTypeLabel: WEBRESOURCE_TYPE_LABELS[i.webresourcetype - 1],
            isManaged: i.ismanaged,
            isManagedLabel: i.ismanaged ? 'Managed' : 'Unmanaged',
            canBeDeleted: i.canbedeleted.Value
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async checkIfSolutionExists(solutionId: string, dynamicsApiUrl: string, logger: Logger): Promise<boolean> {
    const solutionRequestUrl = `${dynamicsApiUrl}/api/data/v9.2/solutions?$filter=solutionid eq ${solutionId}&$count=true`;
    const solutionResponse = await odata.getAllItems<Solution>(solutionRequestUrl);

    this.logIfVerbose(logger, `Retrieved solution with ID '${solutionId}'`);

    return solutionResponse.length === 1;
  }

  private async logIfVerbose(logger: Logger, message: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(message);
    }
  }

  private getWebresourceProperties(args: CommandArgs): string[] {
    if (cli.shouldTrimOutput(args.options.output)) {
      return this.defaultProperties()!.filter(defaultProperty =>
        defaultProperty !== 'webresourceTypeLabel' && defaultProperty !== 'isManagedLabel'
      )!.map(prop =>
        prop.toLowerCase()
      );
    }
    else if (args.options.excludeContent) {
      return [...this.WEBRESOURCE_PROPS_ALL].filter(prop => ['content', 'content_binary'].includes(prop) === false);
    }
    else {
      return this.WEBRESOURCE_PROPS_ALL;
    }
  }
}

export default new PpDataverseWebResourceListCommand();