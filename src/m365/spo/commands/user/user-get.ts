import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  email?: string;
  id: string | number | undefined;
  loginName?: string;
}

class SpoUserGetCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets a site user within specific web';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.email = (!(!args.options.email)).toString();
    telemetryProps.loginName = (!(!args.options.loginName)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information for list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/siteusers/GetById('${encodeURIComponent(args.options.id as string)}')`;
    }
    else if (args.options.email) {
      requestUrl = `${args.options.webUrl}/_api/web/siteusers/GetByEmail('${encodeURIComponent(args.options.email as string)}')`;
    }
    else if (args.options.loginName) {
      requestUrl = `${args.options.webUrl}/_api/web/siteusers/GetByLoginName('${encodeURIComponent(args.options.loginName as string)}')`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((userInstance): void => {
        cmd.log(userInstance);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the web to list the user within'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the user to retrieve information for. Use either "email", "id" or "loginName", but not all.'
      },
      {
        option: '--email [email]',
        description: 'Email address of user to retrieve information for. Use either "email", "id" or "loginName", but not all.'
      },
      {
        option: '--loginName [loginName]',
        description: 'Login name of the user to retrieve information for. Use either "email", "id" or "loginName", but not all.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.id && !args.options.email && !args.options.loginName) {
        return 'Specify id, email or loginName, one is required';
      }

      if ((args.options.id && args.options.email) ||
        (args.options.id && args.options.loginName) ||
        (args.options.loginName && args.options.email)) {
        return 'Use either email, id or loginName, but not all';
      }

      if (args.options.id &&
        typeof args.options.id !== 'number') {
        return `Specified id ${args.options.id} is not a number`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get user with email address ${chalk.grey('john.doe@mytenant.onmicrosoft.com')} for web
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.USER_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --email john.doe@mytenant.onmicrosoft.com

    Get user with id ${chalk.grey('6')} for web ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.USER_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --id 6

    Get user with login name ${chalk.grey('i:0#.f|membership|john.doe@mytenant.onmicrosoft.com')}
    for web  ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.USER_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --loginName "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
    `);
  }
}

module.exports = new SpoUserGetCommand();
