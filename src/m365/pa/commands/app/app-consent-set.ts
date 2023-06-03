import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string,
  name: string;
  bypass: boolean;
  confirm?: boolean;
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
        option: '-e, --environment <environment>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-b, --bypass <bypass>',
        autocomplete: ['true', 'false']
      },
      {
        option: '--confirm'
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
      logger.logToStderr(`Setting the bypass consent for the Microsoft Power App ${args.options.name}... to ${args.options.bypass}`);
    }

    if (args.options.confirm) {
      await this.consentPaApp(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you bypass the consent for the Microsoft Power App ${args.options.name} to ${args.options.bypass}?`
      });

      if (result.continue) {
        await this.consentPaApp(args);
      }
    }
  }

  private async consentPaApp(args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environment}/apps/${args.options.name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`,
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

module.exports = new PaAppConsentSetCommand();