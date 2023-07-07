import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { spo } from '../../../../utils/spo';
import { formatting } from '../../../../utils/formatting';
import { CustomAction } from '../../commands/customaction/customaction';
import * as os from 'os';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  scope?: string;
  confirm?: boolean;
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
        option: '--confirm'
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
        confirm: !!args.options.confirm
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
      if (args.options.confirm) {
        return await this.removeApplicationCustomizer(logger, args.options);
      }

      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the application customizer '${args.options.clientSideComponentId || args.options.title || args.options.id}'?`
      });

      if (result.continue) {
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
      logger.logToStderr(`Removing application customizer '${options.clientSideComponentId || options.title || options.id}' from the site '${options.webUrl}'...`);
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
      throw `Multiple application customizer with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found. Please disambiguate using IDs: ${os.EOL}${appCustomizers.map(a => `- ${a.Id}`).join(os.EOL)}`;
    }

    return appCustomizers[0];
  }
}

module.exports = new SpoApplicationCustomizerRemoveCommand();