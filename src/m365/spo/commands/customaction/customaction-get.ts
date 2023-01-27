import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from './customaction';
import { Options as SpoCustomActionListCommandOptions } from './customaction-list';
import * as SpoCustomActionListCommand from './customaction-list';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  clientSideComponentId?: string;
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
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
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
        option: '-c, --clientSideComponentId [clientSideComponentId]'
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

        if (args.options.clientSideComponentId && validation.isValidGuid(args.options.clientSideComponentId) === false) {
          return `${args.options.clientSideComponentId} is not valid. Custom action clientSideComponentId (Guid) expected.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
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

  private async getCustomAction(options: Options): Promise<CustomAction> {
    if (options.clientSideComponentId) {
      const customActionListCommandoptions: SpoCustomActionListCommandOptions = {
        webUrl: options.webUrl,
        scope: options.scope,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const output = await Cli.executeCommandWithOutput(SpoCustomActionListCommand as Command, { options: { ...customActionListCommandoptions, _: [] } });

      if (!output.stdout) {
        throw `No user custom action with ClientSideComponentId '${options.clientSideComponentId}' found`;
      }
      const getCustomActionListOutput: CustomAction[] = JSON.parse(output.stdout);
      const result: CustomAction[] = getCustomActionListOutput.filter((x: CustomAction) => x.ClientSideComponentId === options.clientSideComponentId);

      if (result.length === 0) {
        throw `No user custom action with ClientSideComponentId '${options.clientSideComponentId}' found`;
      }

      if (result.length > 1) {
        throw `Multiple user custom actions with ClientSideComponentId '${options.clientSideComponentId}' found. Please disambiguate using IDs: ${result.map(a => a.Id).join(', ')}`;
      }

      return result[0];
    }
    else {
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
        const result = await request.get<CustomAction>(requestOptions);
        return result;
      }

      const result = await request.get<{ value: CustomAction[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `No user custom action with title '${options.title}' found`;
      }

      if (result.value.length > 1) {
        throw `Multiple user custom actions with title '${options.title}' found. Please disambiguate using IDs: ${result.value.map(a => a.Id).join(', ')}`;
      }

      return result.value[0];
    }
  }

  /**
   * Get request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private async searchAllScopes(options: Options): Promise<CustomAction> {
    try {
      options.scope = "Web";

      const webResult = await this.getCustomAction(options);
      if (webResult["odata.null"] !== true) {
        return webResult;
      }

      options.scope = "Site";
      const siteResult = await this.getCustomAction(options);
      return siteResult;

    }
    catch (err) {
      throw err;
    }
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