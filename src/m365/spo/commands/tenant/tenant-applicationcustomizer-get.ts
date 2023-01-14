import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
import * as SpoListItemListCommand from '../listitem/listitem-list';
import { Options as SpoListItemListCommandOptions } from '../listitem/listitem-list';
import { ListItemInstance } from '../listitem/ListItemInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  id?: string;
  clientSideComponentId?: string;
}

class SpoTenantApplicationCustomizerGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_GET;
  }

  public get description(): string {
    return 'Get an application customizer that is installed tenant wide';
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
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title [title]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
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
      const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(SpoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
      const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      let filter: string = '';
      if (args.options.title) {
        filter = `Title eq '${args.options.title}'`;
      }
      else if (args.options.id) {
        filter = `GUID eq '${args.options.id}'`;
      }
      else if (args.options.clientSideComponentId) {
        filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
      }

      const options: SpoListItemListCommandOptions = {
        output: 'json',
        debug: args.options.debug,
        verbose: args.options.verbose,
        listTitle: 'Tenant Wide Extensions',
        webUrl: appCatalogUrl,
        filter: filter
      };

      const spoListItemGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(SpoListItemListCommand as Command, { options: { ...options, _: [] } });
      const listItemOutput = JSON.parse(spoListItemGetCommandOutput.stdout) as ListItemInstance[];

      if (listItemOutput.length > 0) {
        logger.log(listItemOutput[0]);
      }
      else {
        throw `The specified application customizer was not found`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTenantApplicationCustomizerGetCommand();