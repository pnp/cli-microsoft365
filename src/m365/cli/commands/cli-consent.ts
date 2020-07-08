import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandAction
} from '../../../Command';
import config from '../../../config';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  service: string;
}

class CliConsentCommand extends Command {
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let scope = '';
    switch (args.options.service) {
      case 'yammer':
        scope = 'https://api.yammer.com/user_impersonation';
        break;
    }

    cmd.log(`To consent permissions for executing ${args.options.service} commands, navigate in your web browser to https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/authorize?client_id=${config.cliAadAppId}&response_type=code&scope=${encodeURIComponent(scope)}`);
    cb();
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      args = (cmd as any).processArgs(args);
      (cmd as any).initAction(args, this);

      cmd.commandAction(this, args, cb);
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --service <service>',
        description: 'Service for which to consent permissions. Allowed values: yammer',
        autocomplete: ['yammer']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.service) {
        return 'Required option service missing';        
      }

      if (args.options.service !== 'yammer') {
        return `${args.options.service} is not a valid value for the service option. Allowed values: yammer`;
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
    
    Using the ${chalk.blue(this.name)} command you can consent additional permissions
    for the Azure AD application used by the CLI for Microsoft 365. This is for example
    necessary to use Yammer commands, which require the Yammer API permission
    that isn't granted to the CLI by default.

    After executing the command, the CLI for Microsoft 365 will present you with a URL
    that you need to open in the web browser in order to consent the permissions
    for the selected Microsoft 365 service.

    To simplify things, rather than wondering which permissions you should
    grant for which CLI commands, this command allows you to easily grant all
    the necessary permissions for using commands for the specified Microsoft 365
    service, like Yammer.

  Examples:
  
    Consent permissions to the Yammer API
      ${commands.CONSENT} --service yammer
`);
  }
}

module.exports = new CliConsentCommand();