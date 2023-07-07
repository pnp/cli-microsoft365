import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from './customaction';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  webUrl: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCustomActionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified custom action';
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
        scope: args.options.scope || 'All',
        confirm: (!(!args.options.confirm)).toString()
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
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && validation.isValidGuid(args.options.id) === false) {
          return `${args.options.id} is not valid. Custom action Id (GUID) expected.`;
        }

        const isValidUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
          return isValidUrl;
        }

        if (args.options.scope) {
          if (args.options.scope !== 'Site' &&
            args.options.scope !== 'Web' &&
            args.options.scope !== 'All') {
            return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeCustomAction: () => Promise<void> = async (): Promise<void> => {
      try {
        let customAction: CustomAction | void;
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          customAction = await this.removeScopedCustomAction(args.options);
        }
        else {
          customAction = await this.searchAllScopes(args.options);
        }

        if (this.verbose) {
          if (customAction && customAction["odata.null"] === true) {
            logger.logToStderr(`Custom action with id ${args.options.id} not found`);
          }
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeCustomAction();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${args.options.id} user custom action?`
      });

      if (result.continue) {
        await removeCustomAction();
      }
    }
  }

  private async getCustomActionId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const customActions = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title as string)}'`);

    if (customActions.length === 1) {
      return customActions[0].Id;
    }

    if (customActions.length === 0) {
      throw `No user custom action with title '${options.title}' found`;
    }

    throw `Multiple user custom actions with title '${options.title}' found. Please disambiguate using IDs: ${customActions.map(a => a.Id).join(', ')}`;
  }

  private async removeScopedCustomAction(options: Options): Promise<CustomAction | undefined> {
    const customActionId = await this.getCustomActionId(options);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions('${formatting.encodeQueryParameter(customActionId)}')')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE'
      },
      responseType: 'json'
    };

    return await request.post(requestOptions);
  }

  /**
   * Remove request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private async searchAllScopes(options: Options): Promise<CustomAction | undefined> {
    options.scope = "Web";
    const webResult = await this.removeScopedCustomAction(options);
    if (!webResult) {
      return webResult;
    }

    options.scope = "Site";
    const siteResult = await this.removeScopedCustomAction(options);

    return siteResult;
  }
}

module.exports = new SpoCustomActionRemoveCommand();