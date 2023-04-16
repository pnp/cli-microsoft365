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
  private static readonly scopes: string[] = ['Site', 'Web', 'All'];

  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_REMOVE;
  }

  public get description(): string {
    return 'Remove an application customizer from a site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
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
        option: '-s, --scope [scope]', autocomplete: SpoApplicationCustomizerRemoveCommand.scopes
      },
      {
        option: '--confirm'
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(this.validateUrl);
    this.validators.push(this.validateParams);
    this.validators.push(this.validateGUIds);
    this.validators.push(this.validateScope);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { options }: CommandArgs = args;
    const removeAppCustomizer: () => Promise<void> = async (): Promise<void> => {
      try {
        const appCustomizers = await this.getAppCustomizers(args);
        const appCustomizerToRemove: CustomAction = this.getAppCustomizerToRemove(appCustomizers);
        if (appCustomizerToRemove) {
          const scope = appCustomizerToRemove.Scope.toString() === '2' ? 'Site' : 'Web';
          await this.removeAppCustomizer(scope, options.webUrl, appCustomizerToRemove.Id);
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };
    const prompt = await this.confirmPrompt({ options });
    if (prompt) {
      await removeAppCustomizer();
    }
  }

  private async validateUrl({ options }: CommandArgs): Promise<boolean | string> {
    const { webUrl } = options;
    if (webUrl) {
      const isValidSharePointUrl = validation.isValidSharePointUrl(webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }
    }
    return true;
  }

  private async validateParams({ options }: CommandArgs): Promise<boolean | string> {
    const { id, clientSideComponentId, title } = options;
    if (id || clientSideComponentId || title) {
      return true;
    }
    else {
      return `At least one of the parameters id, clientSideComponentId, or title must have a value`;
    }
  }

  private async validateGUIds({ options }: CommandArgs): Promise<boolean | string> {
    const { clientSideComponentId, id } = options;
    if (clientSideComponentId && !validation.isValidGuid(clientSideComponentId)) {
      return `${clientSideComponentId} is not a valid GUID`;
    }

    if (id && !validation.isValidGuid(id)) {
      return `${id} is not a valid GUID`;
    }
    return true;
  }

  private async validateScope(args: CommandArgs): Promise<boolean | string> {
    const { options } = args;
    const { scope } = options;
    if (scope && SpoApplicationCustomizerRemoveCommand.scopes.indexOf(scope) < 0) {
      return `${scope} is not a valid value for scope. Valid values are ${SpoApplicationCustomizerRemoveCommand.scopes.join(', ')}`;
    }
    return true;
  }

  private async getAppCustomizers(args: CommandArgs): Promise<CustomAction[]> {
    const { options }: CommandArgs = args;
    const { id, webUrl, title, clientSideComponentId, scope }: Options = options;
    const appCustomizers = await spo.getCustomActions(webUrl, scope, `Location eq 'ClientSideExtension.ApplicationCustomizer'`);
    const filteredAppCustomizers: CustomAction[] = appCustomizers.filter(appCustomizer =>
      (id && appCustomizer.Id.includes(id)) ||
      (!id && appCustomizer.Title.includes(`${title}`)) ||
      (!id && appCustomizer.ClientSideComponentId.includes(`${clientSideComponentId}`))
    );
    return filteredAppCustomizers;
  }

  private getAppCustomizerToRemove(customActions: CustomAction[]): CustomAction {
    const customActionsCount = customActions.length;
    if (customActionsCount === 0) {
      throw `No application customizer found`;
    }
    if (customActionsCount > 1) {
      const ids = customActions.map(a => a.Id).join(', ');
      throw `Multiple application customizer found. Please disambiguate using IDs: ${ids}`;
    }
    return customActions[0];
  }

  private removeAppCustomizer(scope: string, webUrl: string, id: string): Promise<void> {
    const requestOptions: any = {
      url: `${webUrl}/_api/${scope}/UserCustomActions('${id}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.delete(requestOptions);
  }

  private async confirmPrompt({ options }: CommandArgs): Promise<boolean> {
    let confirmation: boolean = false;
    const { id, title, clientSideComponentId, confirm }: Options = options;
    let v: number | string | undefined;
    if (id !== undefined) {
      v = id;
    }
    else if (title !== undefined) {
      v = title;
    }
    else if (clientSideComponentId !== undefined) {
      v = clientSideComponentId;
    }

    const result: { continue: boolean } = confirm
      ? { continue: true }
      : await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${v} application customizer?`
      });

    if (result.continue) {
      confirmation = true;
    }
    return confirmation;
  }
}

module.exports = new SpoApplicationCustomizerRemoveCommand();