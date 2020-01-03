import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Command, {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName: string;
}

class TenantIdGetCommand extends Command {
  public get name(): string {
    return `${commands.TENANT_ID_GET}`;
  }

  public get description(): string {
    return 'Gets Office 365 tenant ID for the specified domain';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `https://login.windows.net/${args.options.domainName}/.well-known/openid-configuration`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json',
        'x-anonymous': true
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res.error) {
          cb(new CommandError(res.error_description));
          return;
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
        description: 'The domain name for which to retrieve the Office 365 tenant ID'
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
    log(vorpal.find(commands.TENANT_ID_GET).helpInformation());
    log(
      `Examples:
  
    Get Office 365 tenant ID for the specified domain
      ${commands.TENANT_ID_GET} --domainName contoso.com
`);
  }
}

module.exports = new TenantIdGetCommand();