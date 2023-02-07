import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import Command from '../../../../Command';
import { Cli, CommandOutput } from '../../../../cli/Cli';
import { Options as spoListItemAddCommandOptions } from '../listitem/listitem-add';
import { Options as spoListItemListCommandOptions } from '../listitem/listitem-list';
import * as spoTenantAppCatalogUrlGetCommand from '../tenant/tenant-appcatalogurl-get';
import * as spoListItemAddCommand from '../listitem/listitem-add';
import * as spoListItemListCommand from '../listitem/listitem-list';
import { urlUtil } from '../../../../utils/urlUtil';

interface Solution {
  ContainsTenantWideExtension: boolean;
  SkipFeatureDeployment: boolean;
}

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
    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
    if (this.verbose) {
      logger.logToStderr(spoTenantAppCatalogUrlGetCommandOutput.stderr);
    }

    const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
    if (!appCatalogUrl) {
      throw 'Cannot add tenant-wide application customizer as app catalog cannot be found';
    }
    if (this.verbose) {
      logger.logToStderr(`Got tenant app catalog url: ${appCatalogUrl}`);
    }

    return appCatalogUrl;
  }

  private async getComponentManifest(appCatalogUrl: string, clientSideComponentId: string, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr('Retrieving component manifest item from the ComponentManifests list on the app catalog site so that we get the solution id');
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef><FieldRef Name='ClientComponentManifest'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug,
      output: 'json'
    };

    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    if (this.verbose) {
      logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);
    if (outputParsed.length === 0) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }

    return outputParsed[0];
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<Solution> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
    }

    const camlQuery = `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef><FieldRef Name='ContainsTenantWideExtension'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`;
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug,
      output: 'json'
    };

    const output = await Cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    if (this.verbose) {
      logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);
    if (outputParsed.length === 0) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }
    return outputParsed[0];
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
      TenantWideExtensionLocation: 'ClientSideExtension.ApplicationCustomizer',
      TenantWideExtensionSequence: 0,
      TenantWideExtensionListTemplate: 0,
      TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
      TenantWideExtensionWebTemplate: options.webTemplate || '',
      TenantWideExtensionDisabled: false,
      verbose: this.verbose,
      debug: this.debug,
      output: options.output
    };

    await Cli.executeCommand(spoListItemAddCommand as Command, { options: { ...commandOptions, _: [] } });
  }
}

module.exports = new SpoTenantApplicationCustomizerAddCommand();