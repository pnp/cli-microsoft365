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
  email: string;
  group?:string;
}

class SpoUserAddCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_ADD;
  }

  public get description(): string {
    return 'Adds user to a specific web';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.webUrl = (!(!args.options.webUrl)).toString();
    telemetryProps.email = (!(!args.options.email)).toString();
    telemetryProps.group=(!(!args.options.group)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const groupRequestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${encodeURIComponent(args.options.group as string)}')`;
    
    const requestOptions: any = {
      url: groupRequestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
    .get<{ Id: string }>(requestOptions)
      .then((res: { Id: string; }): Promise<{}> => {
        const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/GetById('${res.Id}')/users`;

        const requestOptions: any = {
          url: requestUrl,
          headers: {
            "Accept": "application/json;odata=verbose",
          },
          json: true,
          body: {
            '__metadata': {
              'type': 'SP.User'
            },
            'LoginName': `i:0#.f|membership|${args.options.email}`
          }
        };
        return request.post(requestOptions);
      })
      .then((): void => {
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the web to list the user within'
      },
      {
        option: '--email [email]',
        description: 'Email address of user to retrieve information for. Use either "email", "id" or "loginName", but not all.'
      },
      {
        option:'--group [GroupName]',
        description:'Group name from the sites to ad user with in'
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

      if(!args.options.email){
        return 'Required parameter email missing';
      }

      if(!args.options.group){
        return 'Required parameter group name missing';
      }
      
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
    Adds user with email ${chalk.grey('john.doe@contoso.onmicrosoft.com')} 
    to the Viewers group of web ${chalk.grey('https://contoso.sharepoint.com/sites/mysite')} 

    ${commands.USER_ADD} --webUrl "https://contoso.sharepoint.com/sites/mysite" --email "john.doe@contoso.onmicrosoft.com" --group "Team Site Members"
    `);
  }
}

module.exports = new SpoUserAddCommand();
