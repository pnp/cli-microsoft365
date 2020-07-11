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
}

class SpoUserListCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return 'Lists all the users within specific web';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving users from web ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    requestUrl = `${args.options.webUrl}/_api/web/siteusers`;

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
      .then((users: any): void => {
        if (args.options.output === 'json') {
          cmd.log(users);
        }
        else {
          cmd.log(users.value.map((user: any) => {
            return {
              Id: user.Id,
              Title:user.Title,
              LoginName: user.LoginName
            };
          }));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the web to list the users from'
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
      
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Get list of users in web ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.USER_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x 
    `);
  }
}

module.exports = new SpoUserListCommand();