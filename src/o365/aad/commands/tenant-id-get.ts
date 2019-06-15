import config from '../../../config';
import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import Command, {
  CommandOption,
  CommandValidate
} from '../../../Command';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName: string;
}

class AadTenantGetIDCommand extends Command {

  public get name(): string {
    return `${commands.TENANT_ID_GET}`;
  }

  public get description(): string {
    return 'Gets Microsoft Azure or Office 365 tenant ID';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.domainName = typeof args.options.domainName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {

    const endpoint: string = `https://login.windows.net/${args.options.domainName}/.well-known/openid-configuration`;

    const requestOptions: any = {
      url: endpoint,
      headers: {
        "content-type": "application/json",
        accept: 'application/json'
      },
      json: true,
    };

    request.get(requestOptions)
      .then((res: any): void => {

        if (res.error) {
          cmd.log(res.error_description);
        }

        if (res.token_endpoint) {
          cmd.log(res.token_endpoint.split('/')[3]);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-d, --domainName <domainName>',
        description: 'The domain name to get the Microsoft Azure or Office 365 tenant ID'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.domainName) {
        return 'Required parameter domainName missing';
      }

      return true;
    };
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.TENANT_ID_GET).helpInformation());
    log(
      `
  Examples:
  
  Gets Microsoft Azure or Office 365 tenant ID 
      ${chalk.grey(config.delimiter)} ${commands.TENANT_ID_GET} --domainName contoso.com
`);
  }

}

module.exports = new AadTenantGetIDCommand();