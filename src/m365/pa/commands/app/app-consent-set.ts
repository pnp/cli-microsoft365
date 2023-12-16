import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string,
  name: string;
  bypass: boolean;
  force?: boolean;
}

class PaAppConsentSetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_CONSENT_SET;
  }

  public get description(): string {
    return 'Configures if users can bypass the API Consent window for the selected canvas app';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-b, --bypass <bypass>',
        autocomplete: ['true', 'false']
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('bypass');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Setting the bypass consent for the Microsoft Power App ${args.options.name}... to ${args.options.bypass}`);
    }

    if (args.options.force) {
      await this.consentPaApp(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you bypass the consent for the Microsoft Power App ${args.options.name} to ${args.options.bypass}?` });

      if (result) {
        await this.consentPaApp(args);
      }
    }
  }

  private async consentPaApp(args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        bypassconsent: args.options.bypass
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaAppConsentSetCommand();