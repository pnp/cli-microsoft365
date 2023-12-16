import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CustomAction } from './customaction.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  webUrl: string;
  scope?: string;
  clientSideComponentId?: string;
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
          return `${args.options.clientSideComponentId} is not a valid GUID.`;
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
      const customAction = await this.getCustomAction(args.options);

      if (customAction) {
        await logger.log({
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

  private async getCustomAction(options: Options): Promise<CustomAction | undefined> {
    if (options.id) {
      const customAction: CustomAction | undefined = await spo.getCustomActionById(options.webUrl, options.id, options.scope);

      if (!customAction) {
        throw `No user custom action with id '${options.id}' found`;
      }

      return customAction;
    }
    else if (options.title) {
      const customActions: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `Title eq '${formatting.encodeQueryParameter(options.title as string)}'`);

      if (customActions.length === 1) {
        return customActions[0];
      }

      if (customActions.length === 0) {
        throw `No user custom action with title '${options.title}' found`;
      }

      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple user custom actions with title '${options.title}' found.`, resultAsKeyValuePair);
    }
    else {
      const customActions: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `ClientSideComponentId eq guid'${options.clientSideComponentId}'`);

      if (customActions.length === 0) {
        throw `No user custom action with ClientSideComponentId '${options.clientSideComponentId}' found`;
      }

      if (customActions.length > 1) {
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
        return await cli.handleMultipleResultsFound<CustomAction>(`Multiple user custom actions with ClientSideComponentId '${options.clientSideComponentId}' found.`, resultAsKeyValuePair);
      }

      return customActions[0];
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

export default new SpoCustomActionGetCommand();