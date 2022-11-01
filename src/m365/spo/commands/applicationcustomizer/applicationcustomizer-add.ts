import { Cli, CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Options as spoCustomActionAddCommandOptions } from '../customaction/customaction-add';
import { Options as spoListItemAddCommandOptions } from '../listitem/listitem-add';
import { Options as spoListItemListCommandOptions } from '../listitem/listitem-list';
import * as spoTenantAppCatalogUrlGetCommand from '../tenant/tenant-appcatalogurl-get';
import * as spoListItemAddCommand from '../listitem/listitem-add';
import * as spoListItemListCommand from '../listitem/listitem-list';
import * as spoCustomActionAddCommand from '../customaction/customaction-add';
import { urlUtil } from '../../../../utils/urlUtil';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl?: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
}

class SpoApplicationCustomizerAddCommand extends SpoCommand {
  private location = 'ClientSideExtension.ApplicationCustomizer';
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_ADD;
  }

  public get description(): string {
    return 'Add an application customizer to a site';
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
        webTemplate: typeof args.options.webTemplate !== 'undefined',
        webUrl: typeof args.options.webUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-u, --webUrl [webUrl]'
      },
      {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
      },
      {
        option: '--clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '--webTemplate [webTemplate]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.webUrl) {
          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (args.options.webUrl && args.options.webTemplate) {
          return `The options 'webUrl' and 'webTemplate' cannot be set together`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Adding application customizer with ${args.options.clientSideComponentId} to ${args.options.webUrl ? 'specific web' : 'tenant app catalog'}`);
      }

      if (args.options.webUrl) {
        const options: spoCustomActionAddCommandOptions = {
          webUrl: args.options.webUrl,
          name: args.options.title,
          title: args.options.title,
          clientSideComponentId: args.options.clientSideComponentId,
          clientSideComponentProperties: args.options.clientSideComponentProperties || '',
          location: this.location,
          debug: this.debug,
          verbose: this.verbose
        };
        await Cli.executeCommand(spoCustomActionAddCommand as Command, { options: { ...options, _: [] } });
      }
      else {
        const appCatalogUrl = await this.getAppCatalogUrl();
        if (this.verbose) {
          logger.logToStderr(`Got tenant app catalog url: ${appCatalogUrl}`);
        }
        const componentIdExists = await this.checkIfComponentIdExists(appCatalogUrl, args.options, logger);
        if (!componentIdExists) {
          throw 'The solution has not been deployed to all sites. Make sure to deploy this solution to all sites';
        }
        else {
          await this.addTenantWideExtension(appCatalogUrl, args.options, logger);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }

  private async getAppCatalogUrl(): Promise<string> {
    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
    const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
    if (!appCatalogUrl) {
      throw 'Cannot add tenant-wide application customizer as app catalog cannot be found';
    }
    return appCatalogUrl;
  }

  private async checkIfComponentIdExists(appCatalogUrl: string, options: Options, logger: Logger): Promise<boolean> {
    const solutionId = await this.getSolutionIdFromComponentManifestItem(appCatalogUrl, options.clientSideComponentId, logger);
    const skipFeatureDeployment = await this.getSolutionFromAppCatalog(appCatalogUrl, solutionId, logger);
    return skipFeatureDeployment;
  }

  private async getSolutionIdFromComponentManifestItem(appCatalogUrl: string, clientSideComponentId: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr('Retrieving component manifest item from the ComponentManifests list on the app catalog site so that we get the solution id');
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug
    };
    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    const outputParsed = JSON.parse(output.stdout);
    if (!outputParsed.length) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }
    if (outputParsed.length > 1) {
      throw 'Multiple components found with the specified clientSideComponentId. Make sure that this is unique';
    }
    return outputParsed[0].SolutionId;
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<boolean> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug
    };
    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    const outputParsed = JSON.parse(output.stdout);
    if (!outputParsed.length) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }
    if (outputParsed.length > 1) {
      throw `Multiple components found with the solution id ${solutionId}. Make sure that this is unique`;
    }
    return outputParsed[0].SkipFeatureDeployment;
  }

  private async addTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Pre-checks finished. Adding tenant wide extension to the TenantWideExtensions list');
    }

    const commandOptions: spoListItemAddCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`,
      Title: options.title,
      TenantWideExtensionComponentId: options.clientSideComponentId,
      TenantWideExtensionLocation: this.location,
      TenantWideExtensionSequence: 0,
      TenantWideExtensionListTemplate: 0,
      TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
      TenantWideExtensionWebTemplate: options.webTemplate || '',
      TenantWideExtensionDisabled: false,
      verbose: this.verbose,
      debug: this.debug
    };
    await Cli.executeCommand(spoListItemAddCommand as Command, { options: { ...commandOptions, _: [] } });
  }
}

module.exports = new SpoApplicationCustomizerAddCommand();