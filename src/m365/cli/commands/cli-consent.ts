import { Logger } from '../../../cli/Logger.js';
import config from '../../../config.js';
import GlobalOptions from '../../../GlobalOptions.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service: string;
}

class CliConsentCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONSENT;
  }

  public get description(): string {
    return 'Consent additional permissions for the Azure AD application used by the CLI for Microsoft 365';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        service: args.options.service
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --service <service>',
        autocomplete: ['"viva engage"']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.service !== 'viva engage') {
          return `${args.options.service} is not a valid value for the service option. Allowed values: "viva engage"`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let scope = '';
    switch (args.options.service) {
      case 'viva engage':
        scope = 'https://api.yammer.com/user_impersonation';
        break;
    }

    await logger.log(`To consent permissions for executing ${args.options.service} commands, navigate in your web browser to https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/authorize?client_id=${config.cliAadAppId}&response_type=code&scope=${encodeURIComponent(scope)}`);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

export default new CliConsentCommand();