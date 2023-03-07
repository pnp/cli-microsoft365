import { Logger } from '../../../../cli/Logger';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from '../customaction/customaction';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  scope: string;
}

class SpoApplicationCustomizerGetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_GET;
  }

  public get description(): string {
    return 'Get an application customizer that is added to a site.';
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
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined'
      });
    });
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
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      },
      {
        option: '-s, --scope [scope]'
      },
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

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
      const customAction = await this.getCustomAction(args.options);

      if (customAction) {
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

  private async getCustomAction(options: Options): Promise<CustomAction | undefined> {
    const scope = options.scope ? options.scope : 'All';

    if (options.id) {
      const customAction = await spo.getCustomActionById(options.webUrl, options.id, scope);

      if (!customAction) {
        throw `No application customizer with id '${options.id}' found`;
      }

      if (customAction.Location !== 'ClientSideExtension.ApplicationCustomizer') {
        throw 'The found custom action is not an Application Customizer';
      }

      return customAction;
    }

    const filter = options.title ? `Title eq '${formatting.encodeQueryParameter(options.title as string)}'` : `ClientSideComponentId eq guid'${formatting.encodeQueryParameter(options.clientSideComponentId as string)}'`;
    const customActions = await spo.getCustomActions(options.webUrl, scope, filter);

    if (customActions.length === 1) {
      if (customActions[0].Location !== 'ClientSideExtension.ApplicationCustomizer') {
        throw 'The found custom action is not an Application Customizer';
      }

      return customActions[0];
    }

    const errorMessage = options.title ? `title '${options.title}'` : `Client Side Component Id '${options.clientSideComponentId}'`;

    if (customActions.length === 0) {
      throw `No application customizer with ${errorMessage} found`;
    }

    throw `Multiple application customizers with ${errorMessage} found. Please disambiguate using IDs: ${customActions.map(a => a.Id).join(', ')}`;
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

module.exports = new SpoApplicationCustomizerGetCommand();