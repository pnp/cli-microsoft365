import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
import { CustomAction } from '../../commands/customaction/customaction.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  scope?: string;
  force?: boolean;
}

class SpoApplicationCustomizerRemoveCommand extends SpoCommand {
  private readonly allowedScopes: string[] = ['Site', 'Web', 'All'];

  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_REMOVE;
  }

  public get description(): string {
    return 'Removes an application customizer that is added to a site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: this.allowedScopes
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        force: !!args.options.force
      });
    });
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

        if (args.options.scope && this.allowedScopes.indexOf(args.options.scope) === -1) {
          return `'${args.options.scope}' is not a valid application customizer scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title', 'clientSideComponentId'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        return await this.removeApplicationCustomizer(logger, args.options);
      }

      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the application customizer '${args.options.clientSideComponentId || args.options.title || args.options.id}'?` });

      if (result) {
        await this.removeApplicationCustomizer(logger, args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeApplicationCustomizer(logger: Logger, options: Options): Promise<void> {
    const applicationCustomizer = await this.getApplicationCustomizer(options);

    if (this.verbose) {
      await logger.logToStderr(`Removing application customizer '${options.clientSideComponentId || options.title || options.id}' from the site '${options.webUrl}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${applicationCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${applicationCustomizer.Id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.delete(requestOptions);
  }

  private async getApplicationCustomizer(options: Options): Promise<CustomAction> {
    const resolvedScope = options.scope || 'All';
    let appCustomizers: CustomAction[] = [];

    if (options.id) {
      const appCustomizer = await spo.getCustomActionById(options.webUrl, options.id, resolvedScope);

      if (appCustomizer) {
        appCustomizers.push(appCustomizer);
      }
    }
    else if (options.title) {
      appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(options.title as string)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }
    else {
      appCustomizers = await spo.getCustomActions(options.webUrl, resolvedScope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }

    if (appCustomizers.length === 0) {
      throw `No application customizer with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
    }

    if (appCustomizers.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', appCustomizers);
      return await Cli.handleMultipleResultsFound<CustomAction>(`Multiple application customizer with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found.`, resultAsKeyValuePair);
    }

    return appCustomizers[0];
  }
}

export default new SpoApplicationCustomizerRemoveCommand();