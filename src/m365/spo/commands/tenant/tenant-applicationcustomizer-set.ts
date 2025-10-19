import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';
import spoListItemListCommand, { Options as spoListItemListCommandOptions } from '../listitem/listitem-list.js';
import { Solution } from './Solution.js';

interface CommandArgs {
  options: Options;
}

interface FormValue {
  FieldName: string;
  FieldValue: string;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  clientSideComponentId?: string;
  newTitle?: string;
  newClientSideComponentId?: string;
  clientSideComponentProperties?: string;
  webTemplate?: string;
}

class SpoTenantApplicationCustomizerSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_SET;
  }

  public get description(): string {
    return 'Updates an Application Customizer that is deployed as a tenant-wide extension';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        newTitle: typeof args.options.newTitle !== 'undefined',
        newClientSideComponentId: typeof args.options.newClientSideComponentId !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        webTemplate: typeof args.options.webTemplate !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      },
      {
        option: '--newTitle [newTitle]'
      },
      {
        option: '--newClientSideComponentId [newClientSideComponentId]'
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
        if (args.options.id && isNaN(parseInt(args.options.id))) {
          return `${args.options.id} is not a number`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.newClientSideComponentId && !validation.isValidGuid(args.options.newClientSideComponentId)) {
          return `${args.options.newClientSideComponentId} is not a valid GUID`;
        }

        if (!args.options.newTitle && !args.options.newClientSideComponentId && !args.options.clientSideComponentProperties && !args.options.webTemplate) {
          return `Please specify an option to be updated`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      if (args.options.newClientSideComponentId !== undefined) {
        const componentManifest = await this.getComponentManifest(appCatalogUrl, args.options.newClientSideComponentId, logger);
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
      }

      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
      const listItemId: number = await this.getListItemId(appCatalogUrl, args.options, listServerRelativeUrl, logger);
      await this.updateTenantWideExtension(appCatalogUrl, args.options, listServerRelativeUrl, listItemId, logger);
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListItemId(appCatalogUrl: string, options: Options, listServerRelativeUrl: string, logger: Logger): Promise<number> {
    const { title, id, clientSideComponentId } = options;
    const filter = title ? `Title eq '${title}'` : id ? `Id eq '${id}'` : `TenantWideExtensionComponentId eq '${clientSideComponentId}'`;

    if (this.verbose) {
      await logger.logToStderr(`Getting tenant-wide application customizer: "${title || id || clientSideComponentId}"...`);
    }

    const listItemInstances = await odata.getAllItems<ListItemInstance>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`);

    if (!listItemInstances || listItemInstances.length === 0) {
      throw 'The specified application customizer was not found';
    }

    if (listItemInstances.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
      const result = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple application customizers with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found.`, resultAsKeyValuePair);
      return result.Id;
    }

    return listItemInstances[0].Id;
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

  private async updateTenantWideExtension(appCatalogUrl: string, options: Options, listServerRelativeUrl: string, itemId: number, logger: Logger): Promise<void> {
    const { title, id, clientSideComponentId, newTitle, newClientSideComponentId, clientSideComponentProperties, webTemplate } = options;

    if (this.verbose) {
      await logger.logToStderr(`Updating tenant-wide application customizer: "${title || id || clientSideComponentId}"...`);
    }

    const formValues: FormValue[] = [];

    if (newTitle !== undefined) {
      formValues.push({
        FieldName: 'Title',
        FieldValue: newTitle
      });
    }

    if (newClientSideComponentId !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentId',
        FieldValue: newClientSideComponentId
      });
    }

    if (clientSideComponentProperties !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentProperties',
        FieldValue: clientSideComponentProperties
      });
    }

    if (webTemplate !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionWebTemplate',
        FieldValue: webTemplate
      });
    }

    const requestOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${itemId})/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        formValues: formValues
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

export default new SpoTenantApplicationCustomizerSetCommand();
