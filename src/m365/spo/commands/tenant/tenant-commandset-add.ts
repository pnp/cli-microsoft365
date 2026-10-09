import { z } from 'zod';
import Command from '../../../../Command.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandOutput, cli } from '../../../../cli/cli.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemAddCommand, { Options as spoListItemAddCommandOptions } from '../listitem/listitem-add.js';
import spoListItemListCommand, { Options as spoListItemListCommandOptions } from '../listitem/listitem-list.js';
import { Solution } from './Solution.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';

const listTypes = ['List', 'Library', 'SitePages'] as const;
const locations = ['ContextMenu', 'CommandBar', 'Both'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().alias('t'),
  listType: z.enum(listTypes).alias('l'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), { message: 'clientSideComponentId is not a valid GUID' }).alias('i'),
  clientSideComponentProperties: z.string().optional().alias('p'),
  webTemplate: z.string().optional().alias('w'),
  location: z.enum(locations).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantCommandSetAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_ADD;
  }

  public get description(): string {
    return 'Adds a ListView Command Set as a tenant-wide extension.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
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
    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
    if (this.verbose) {
      await logger.logToStderr(spoTenantAppCatalogUrlGetCommandOutput.stderr);
    }

    const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
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
    const commandOptions: spoListItemListCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
      camlQuery: camlQuery,
      verbose: this.verbose,
      debug: this.debug,
      output: 'json'
    };

    const output = await cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    if (this.verbose) {
      await logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);
    if (outputParsed.length === 0) {
      throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
    }

    return outputParsed[0];
  }

  private async getSolutionFromAppCatalog(appCatalogUrl: string, solutionId: string, logger: Logger): Promise<Solution> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
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

    const output = await cli.executeCommandWithOutput(spoListItemListCommand as Command, { options: { ...commandOptions, _: [] } });
    if (this.verbose) {
      await logger.logToStderr(output.stderr);
    }

    const outputParsed = JSON.parse(output.stdout);
    if (outputParsed.length === 0) {
      throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
    }

    return outputParsed[0];
  }

  private async addTenantWideExtension(appCatalogUrl: string, options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Pre-checks finished. Adding tenant wide extension to the TenantWideExtensions list');
    }

    const commandOptions: spoListItemAddCommandOptions = {
      webUrl: appCatalogUrl,
      listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`,
      Title: options.title,
      TenantWideExtensionComponentId: options.clientSideComponentId,
      TenantWideExtensionLocation: this.getLocation(options.location),
      TenantWideExtensionSequence: 0,
      TenantWideExtensionListTemplate: this.getListTemplate(options.listType),
      TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
      TenantWideExtensionWebTemplate: options.webTemplate || '',
      TenantWideExtensionDisabled: false,
      verbose: this.verbose,
      debug: this.debug,
      output: options.output
    };

    await cli.executeCommand(spoListItemAddCommand as Command, { options: { ...commandOptions, _: [] } });
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