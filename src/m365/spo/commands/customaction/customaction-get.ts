import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
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
}

class SpoCustomActionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_GET;
  }

  public get description(): string {
    return 'Gets details for the specified custom action';
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
        scope: args.options.scope || 'All'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && validation.isValidGuid(args.options.id) === false) {
          return `${args.options.id} is not valid. Custom action id (Guid) expected.`;
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
    try {
      let customAction: CustomAction;
      if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
        customAction = await this.getCustomAction(args.options);
      }
      else {
        customAction = await this.searchAllScopes(args.options);
      }


      if (customAction["odata.null"] === true) {
        if (this.verbose) {
          logger.logToStderr(`Custom action with id ${args.options.id} not found`);
        }
      }
      else {
        logger.log({
          ClientSideComponentId: customAction.ClientSideComponentId,
          ClientSideComponentProperties: customAction.ClientSideComponentProperties,
          CommandUIExtension: customAction.CommandUIExtension,
          Description: customAction.Description,
          Group: customAction.Group,
          Id: customAction.Id,
          ImageUrl: customAction.ImageUrl,
          Location: customAction.Location,
          Name: customAction.Name,
          RegistrationId: customAction.RegistrationId,
          RegistrationType: customAction.RegistrationType,
          Rights: JSON.stringify(customAction.Rights),
          Scope: this.humanizeScope(customAction.Scope),
          ScriptBlock: customAction.ScriptBlock,
          ScriptSrc: customAction.ScriptSrc,
          Sequence: customAction.Sequence,
          Title: customAction.Title,
          Url: customAction.Url,
          VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
        });
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private getCustomAction(options: Options): Promise<CustomAction> {
    const filter: string = options.id ?
      `('${formatting.encodeQueryParameter(options.id as string)}')` :
      `?$filter=Title eq '${formatting.encodeQueryParameter(options.title as string)}'`;

    const requestOptions: any = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions${filter}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (options.id) {
      return request
        .get<CustomAction>(requestOptions)
        .then((res: CustomAction): Promise<CustomAction> => {
          return Promise.resolve(res);
        });
    }

    return request
      .get<{ value: CustomAction[] }>(requestOptions)
      .then((res: { value: CustomAction[] }): Promise<CustomAction> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0]);
        }

        if (res.value.length === 0) {
          return Promise.reject(`No user custom action with title '${options.title}' found`);
        }

        return Promise.reject(`Multiple user custom actions with title '${options.title}' found. Please disambiguate using IDs: ${res.value.map(a => a.Id).join(', ')}`);
      });
  }

  /**
   * Get request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private searchAllScopes(options: Options): Promise<CustomAction> {
    return new Promise<CustomAction>((resolve: (customAction: CustomAction) => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .getCustomAction(options)
        .then((webResult: CustomAction): void => {
          if (webResult["odata.null"] !== true) {
            return resolve(webResult);
          }

          options.scope = "Site";
          this
            .getCustomAction(options)
            .then((siteResult: CustomAction): void => {
              return resolve(siteResult);
            }, (err: any): void => {
              reject(err);
            });
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }
}

module.exports = new SpoCustomActionGetCommand();