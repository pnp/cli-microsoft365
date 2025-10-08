import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { ListItemAddOptions, ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Solution } from './Solution.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
}

class SpoTenantApplicationCustomizerAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_ADD;
  }

  public get description(): string {
    return 'Add an application customizer as a tenant wide extension.';
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
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-w, --webTemplate [webTemplate]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await this.getAppCatalogUrl(logger);
      const componentManifest = await this.getComponentManifest(appCatalogUrl, args.options.clientSideComponentId, logger);
      const clientComponentManifest = JSON.parse(componentManifest.ClientComponentManifest);

      if (clientComponentManifest.extensionType !== "ApplicationCustomizer") {
        throw `The extension type of this component is not of type 'ApplicationCustomizer' but of type '${clientComponentManifest.extensionType}'`;
      }

      const solution = await this.getSolutionFromAppCatalog(appCatalogUrl, componentManifest.SolutionId, logger);

      if (!solution.ContainsTenantWideExtension) {
        throw `The solution does not contain an extension that can be deployed to all sites. Make sure that you've entered the correct component Id.`;
      }
      else if (!solution.SkipFeatureDeployment) {
        throw 'The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.';
      }

      await this.addTenantWideExtension(appCatalogUrl, args.options, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppCatalogUrl(logger: Logger): Promise<string> {
    const appCatalogUrl: string | null = await spo.getTenantAppCatalogUrl(logger, this.verbose);

    if (!appCatalogUrl) {
      throw 'Cannot add tenant-wide application customizer as app catalog cannot be found';
    }
    if (this.verbose) {
      await logger.logToStderr(`Got tenant app catalog url: ${appCatalogUrl}`);
    }

    return appCatalogUrl;
  }

  private async getComponentManifest(appCatalogUrl: string, clientSideComponentId: string, logger: Logger): Promise<any> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving component manifest item from the ComponentManifests list on the app catalog site so that we get the solution id');
    }

    const options: ListItemListOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
      camlQuery: `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef><FieldRef Name='ClientComponentManifest'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`
    };

    const output = await spoListItem.getListItems(options, logger, this.verbose);

    if (output.length === 0) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }

    return output[0];
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<Solution> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
    }

    const options: ListItemListOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`,
      camlQuery: `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef><FieldRef Name='ContainsTenantWideExtension'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`
    };

    const output = await spoListItem.getListItems(options, logger, this.verbose) as any[];

    if (output.length === 0) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }

    return output[0];
  }

  private async addTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Pre-checks finished. Adding tenant wide extension to the TenantWideExtensions list');
    }

    const listItemAddOptions: ListItemAddOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`,
      fieldValues: {
        Title: options.title,
        TenantWideExtensionComponentId: options.clientSideComponentId,
        TenantWideExtensionLocation: 'ClientSideExtension.ApplicationCustomizer',
        TenantWideExtensionSequence: 0,
        TenantWideExtensionListTemplate: 0,
        TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
        TenantWideExtensionWebTemplate: options.webTemplate || '',
        TenantWideExtensionDisabled: false
      }
    };

    await spoListItem.addListItem(listItemAddOptions, logger, this.verbose, this.debug);
  }
}

export default new SpoTenantApplicationCustomizerAddCommand();