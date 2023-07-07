import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
  newTitle?: string;
  clientSideComponentProperties?: string;
  scope?: string;
}

class SpoApplicationCustomizerSetCommand extends SpoCommand {
  private readonly allowedScopes: string[] = ['Site', 'Web', 'All'];

  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_SET;
  }

  public get description(): string {
    return 'Updates an existing Application Customizer on a site';
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
        option: '--newTitle [newTitle]'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: this.allowedScopes
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        newTitle: typeof args.options.newTitle !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
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

        if (!args.options.newTitle && !args.options.clientSideComponentProperties) {
          return `Please specify an option to be updated`;
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
      const appCustomizer = await this.getAppCustomizerToUpdate(logger, args.options);
      await this.updateAppCustomizer(logger, args.options, appCustomizer);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async updateAppCustomizer(logger: Logger, options: Options, appCustomizer: CustomAction): Promise<void> {
    const { clientSideComponentProperties, webUrl, newTitle }: Options = options;

    if (this.verbose) {
      logger.logToStderr(`Updating application customizer with ID '${appCustomizer.Id}' on the site '${webUrl}'...`);
    }

    const requestBody: any = {};

    if (newTitle) {
      requestBody.Title = newTitle;
    }

    if (clientSideComponentProperties !== undefined) {
      requestBody.ClientSideComponentProperties = clientSideComponentProperties;
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/${appCustomizer.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${appCustomizer.Id}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post<CustomAction>(requestOptions);
  }

  private async getAppCustomizerToUpdate(logger: Logger, options: Options): Promise<CustomAction> {
    const { id, webUrl, title, clientSideComponentId, scope }: Options = options;
    const resolvedScope = scope || 'All';

    if (this.verbose) {
      logger.logToStderr(`Getting application customizer ${title || clientSideComponentId || id} to update...`);
    }

    let appCustomizers: CustomAction[] = [];

    if (id) {
      const appCustomizer = await spo.getCustomActionById(webUrl, id, resolvedScope);
      if (appCustomizer && appCustomizer.Location === 'ClientSideExtension.ApplicationCustomizer') {
        appCustomizers.push(appCustomizer);
      }
    }
    else if (title) {
      appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(Title eq '${formatting.encodeQueryParameter(title as string)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }
    else {
      appCustomizers = await spo.getCustomActions(webUrl, resolvedScope, `(ClientSideComponentId eq guid'${clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`);
    }

    if (appCustomizers.length === 0) {
      throw `No application customizer with ${title && `title '${title}'` || clientSideComponentId && `ClientSideComponentId '${clientSideComponentId}'` || id && `id '${id}'`} found`;
    }

    if (appCustomizers.length > 1) {
      throw `Multiple application customizer with ${title ? `title '${title}'` : `ClientSideComponentId '${clientSideComponentId}'`} found. Please disambiguate using IDs: ${os.EOL}${appCustomizers.map(a => `- ${a.Id}`).join(os.EOL)}`;
    }

    return appCustomizers[0];
  }
}

module.exports = new SpoApplicationCustomizerSetCommand();