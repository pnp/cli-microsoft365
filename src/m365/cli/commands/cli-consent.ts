import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import config from '../../../config';
import GlobalOptions from '../../../GlobalOptions';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service: string;
}

class CliConsentCommand extends AnonymousCommand {
  public get name(): string {
    return `${commands.CONSENT}`;
  }

  public get description(): string {
    return 'Consent additional permissions for the Azure AD application used by the CLI for Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.service = args.options.service;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let scope = '';
    switch (args.options.service) {
      case 'yammer':
        scope = 'https://api.yammer.com/user_impersonation';
        break;
    }

    logger.log(`To consent permissions for executing ${args.options.service} commands, navigate in your web browser to https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/authorize?client_id=${config.cliAadAppId}&response_type=code&scope=${encodeURIComponent(scope)}`);
    cb();
  }

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.initAction(args, logger);
    this.commandAction(logger, args, cb);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --service <service>',
        autocomplete: ['yammer']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.service !== 'yammer') {
      return `${args.options.service} is not a valid value for the service option. Allowed values: yammer`;
    }

    return true;
  }
}

module.exports = new CliConsentCommand();