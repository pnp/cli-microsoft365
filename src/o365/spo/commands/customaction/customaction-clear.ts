import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCustomActionClearCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_CLEAR}`;
  }

  public get description(): string {
    return 'Deletes all custom actions in the collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clearCustomActions = (): void => {
      const resource: string = Auth.getResourceFromUrl(args.options.url);
      let siteAccessToken: string = '';

      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise | Promise<void> => {
          siteAccessToken = accessToken;

          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Clearing custom actions in scope ${args.options.scope}...`);
          }

          if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
            return this.clearScopedCustomActions(args.options, siteAccessToken, cmd);
          }

          return this.clearAllScopes(args.options, siteAccessToken, cmd);
        })
        .then((response: any): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(JSON.stringify(response));
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      clearCustomActions();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear all the user custom actions with scope ${vorpal.chalk.yellow(args.options.scope || 'All')}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          clearCustomActions();
        }
      });
    }
  }

  private clearScopedCustomActions(options: Options, siteAccessToken: string, cmd: CommandInstance): request.RequestPromise {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions/clear`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (this.debug) {
      cmd.log('Executing web request...');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return request.post(requestOptions);
  }

  /**
   * Clear request with `web` scope is send first. 
   * Another clear request is send with `site` scope after.
   */
  private clearAllScopes(options: Options, siteAccessToken: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .clearScopedCustomActions(options, siteAccessToken, cmd)
        .then((webResult: any): request.RequestPromise => {
          if (this.debug) {
            cmd.log('clearScopedCustomActions with scope of web result...');
            cmd.log(JSON.stringify(webResult));
            cmd.log('');
          }

          options.scope = "Site";
          return this.clearScopedCustomActions(options, siteAccessToken, cmd);
        })
        .then((siteResult: any): void => {
          if (this.debug) {
            cmd.log('clearScopedCustomActions with scope of site result...');
            cmd.log(JSON.stringify(siteResult));
            cmd.log('');
          }

          return resolve();
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'Url of the site or site collection to clear the custom actions from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing all custom actions'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Missing required option url';
      }

      const isValidUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (typeof isValidUrl === 'string') {
        return isValidUrl;
      }

      if (args.options.scope &&
        args.options.scope !== 'Site' &&
        args.options.scope !== 'Web' &&
        args.options.scope !== 'All') {
        return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CUSTOMACTION_CLEAR).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
                      
  Remarks:

    To clear user custom actions, you have to first connect to a SharePoint
    Online site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:
  
    Clears all user custom actions for both site and site collection
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}. Skips the confirmation prompt
    message.
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_CLEAR} -u https://contoso.sharepoint.com/sites/test --confirm

    Clears all user custom actions for site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_CLEAR} -u https://contoso.sharepoint.com/sites/test -s Web

    Clears all user custom actions for site collection
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_CLEAR} --url https://contoso.sharepoint.com/sites/test --scope Site
    `);
  }
}

module.exports = new SpoCustomActionClearCommand();