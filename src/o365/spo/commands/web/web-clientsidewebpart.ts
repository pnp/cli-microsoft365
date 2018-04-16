import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebClientSideWebPart extends SpoCommand {
  
  public get name(): string {
    return commands.WEB_CLIENTSIDEWEBPART;
  }

  public get description(): string {
    return 'Lists available client-side web parts';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
    .ensureAccessToken(auth.service.resource, cmd, this.debug)
    .then((): request.RequestPromise => {
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/GetClientSideWebParts`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${auth.service.accessToken}`,
          accept: 'application/json;odata.metadata=none'
        }),
        json: true
      };

      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      return request.get(requestOptions);
    })
    .then((res: any): void => {
      if (this.debug) {
        cmd.log('Response:')
        cmd.log(res);
        cmd.log('');
      }

      cb();
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Web full url'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      if (!args.options.webUrl) {
        const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        return 'webUrl is not a valid SharePoint Url';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.

    Remarks:
    
      To create a subsite, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
    
    Examples:
    
    ` );
  }
}

module.exports = new SpoWebClientSideWebPart();