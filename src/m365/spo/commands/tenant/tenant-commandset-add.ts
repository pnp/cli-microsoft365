import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { spo } from '../../../../utils/spo.js';
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
  listType: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
  location?: string;
}

class SpoTenantCommandSetAddCommand extends SpoCommand {
  private static readonly listTypes: string[] = ['List', 'Library', 'SitePages'];
  private static readonly locations: string[] = ['ContextMenu', 'CommandBar', 'Both'];

  public get name(): string {
    return commands.TENANT_COMMANDSET_ADD;
  }

  public get description(): string {
    return 'Add a ListView Command Set as a tenant-wide extension.';
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
        listType: args.options.listType,
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined',
        location: args.options.location
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-l, --listType <listType>',
        autocomplete: SpoTenantCommandSetAddCommand.listTypes
      },
      {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-w, --webTemplate [webTemplate]'
      },
      {
        option: '--location [location]',
        autocomplete: SpoTenantCommandSetAddCommand.locations
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (SpoTenantCommandSetAddCommand.listTypes.indexOf(args.options.listType) < 0) {
          return `${args.options.listType} is not a valid list type. Allowed values are ${SpoTenantCommandSetAddCommand.listTypes.join(', ')}`;
        }

        if (args.options.location && SpoTenantCommandSetAddCommand.locations.indexOf(args.options.location) < 0) {
          return `${args.options.location} is not a valid location. Allowed values are ${SpoTenantCommandSetAddCommand.locations.join(', ')}`;
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

      if (clientComponentManifest.extensionType !== "ListViewCommandSet") {
        throw `The extension type of this component is not of type 'ListViewCommandSet' but of type '${clientComponentManifest.extensionType}'`;
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
      throw 'Cannot add tenant-wide ListView Command Set as app catalog cannot be found';
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

    const camlQuery = `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef><FieldRef Name='ClientComponentManifest'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`;

    const output = await spo.getListItems(appCatalogUrl, undefined, `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`, camlQuery, undefined, undefined, logger, this.verbose);

    if (this.verbose) {
      await logger.logToStderr(output);
    }

    if (output.length === 0) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }

    return output[0];
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<Solution> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef><FieldRef Name='ContainsTenantWideExtension'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`;

    const output = await spo.getListItems(appCatalogUrl, undefined, `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`, camlQuery, undefined, undefined, logger, this.verbose);

    if (this.verbose) {
      await logger.logToStderr(output);
    }

    if (output.length === 0) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }

    return output[0];
  }

  private async addTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Pre-checks finished. Adding tenant wide extension to the TenantWideExtensions list');
    }

    const listItemAddOptions = {
      Title: options.title,
      TenantWideExtensionComponentId: options.clientSideComponentId,
      TenantWideExtensionLocation: this.getLocation(options.location),
      TenantWideExtensionSequence: 0,
      TenantWideExtensionListTemplate: this.getListTemplate(options.listType),
      TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
      TenantWideExtensionWebTemplate: options.webTemplate || ''
    };

    await spo.addListItem(appCatalogUrl, `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`, listItemAddOptions, logger, this.verbose);
  }

  private getLocation(location: string | undefined): string {
    switch (location) {
      case 'Both':
        return 'ClientSideExtension.ListViewCommandSet';
      case 'ContextMenu':
        return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
      default:
        return 'ClientSideExtension.ListViewCommandSet.CommandBar';
    }
  }

  private getListTemplate(listTemplate: string): string {
    switch (listTemplate) {
      case 'SitePages':
        return '119';
      case 'Library':
        return '101';
      default:
        return '100';
    }
  }
}

export default new SpoTenantCommandSetAddCommand();