import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { spo } from '../../../../utils/spo';
import { CustomAction } from '../../commands/customaction/customaction';

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
    return 'Remove an application customizer that is added to a site';
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
    const { options }: CommandArgs = args;
    const { clientSideComponentId, title, id, webUrl, confirm }: Options = options;

    const customizerIdentifier = clientSideComponentId || title || id;

    if (this.verbose) {
      logger.logToStderr(`Removing application customizer ${customizerIdentifier} from the site '${webUrl}'...`);
    }

    if (confirm) {
      await this.removeAppCustomizer(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the application customizer '${customizerIdentifier}'?`
      });

      if (result.continue) {
        await this.removeAppCustomizer(args);
      }
    }
  }

  private async removeAppCustomizer(args: CommandArgs): Promise<void> {
    const { options }: CommandArgs = args;
    const { scope, webUrl }: Options = options;

    options.scope = scope || 'All';

    try {
      const customAction = await this.getAppCustomizerToRemove(options);

      const requestOptions: any = {
        url: `${webUrl}/_api/${customAction.Scope.toString() === '2' ? 'Site' : 'Web'}/UserCustomActions('${customAction.Id}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      return request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppCustomizerToRemove(options: Options): Promise<CustomAction> {
    const { id, webUrl, title, clientSideComponentId, scope }: Options = options;

    const appCustomizers = await spo.getCustomActions(webUrl, scope, `Location eq 'ClientSideExtension.ApplicationCustomizer'`);

    const filteredAppCustomizers: CustomAction[] = appCustomizers.filter(appCustomizer =>
      (id && appCustomizer.Id.includes(id)) ||
      (!id && appCustomizer.Title.includes(`${title}`)) ||
      (!id && appCustomizer.ClientSideComponentId.includes(`${clientSideComponentId}`))
    ).filter((value, index, self) => {
      return self.findIndex(item => item.Id === value.Id) === index;
    });

    const customActionsCount = filteredAppCustomizers.length;

    if (customActionsCount === 0) {
      throw `No application customizer found`;
    }

    if (customActionsCount > 1) {
      const ids = filteredAppCustomizers.map(a => a.Id).join(', ');
      throw `Multiple application customizer found. Please disambiguate using IDs: ${ids}`;
    }

    return filteredAppCustomizers[0];
  }
}

module.exports = new SpoApplicationCustomizerRemoveCommand();