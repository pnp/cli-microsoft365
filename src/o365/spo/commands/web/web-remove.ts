import auth from '../../SpoAuth';
import * as request from 'request-promise-native';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { Auth } from '../../../../Auth';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  confirm?: boolean;
}

class SpoWebAddCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_REMOVE;
  }

  public get description(): string {
    return 'Delete specified subsite';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeWeb = (): void => {
      const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Deleting subsite ${args.options.webUrl}...`);
          }

          const requestOptions: any = {
            url: `${encodeURI(args.options.webUrl)}/_api/web`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${accessToken}`,
              accept: 'application/json;odata=nometadata',
              'X-HTTP-Method': 'DELETE'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(`Deleting subsite ${args.options.webUrl} ...`);
          }

          return request.post(requestOptions)
        })
        .then((res: any): void => {
          if (this.debug) {
            cmd.log('Response:')
            cmd.log(res.statusCode);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      removeWeb();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the subsite ${args.options.webUrl}`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeWeb();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the subsite to remove'
      },
      {
        option: '--confirm',
        description: 'Do not prompt for confirmation before deleting the subsite'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }

      const isValidUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (typeof isValidUrl === 'string') {
        return isValidUrl;
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
  
    To delete a subsite, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
  
  Examples:
  
    Delete subsite without prompting for confirmation
      ${chalk.grey(config.delimiter)} ${commands.WEB_REMOVE} --webUrl https://contoso.sharepoint.com/subsite --confirm
  ` );
  }
}

module.exports = new SpoWebAddCommand();