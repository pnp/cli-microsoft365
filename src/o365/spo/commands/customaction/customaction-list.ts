import auth from '../../SpoAuth';
import { Auth } from '../../../../Auth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';
import { CustomAction } from './customaction';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_LIST}`;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.url);
    let siteAccessToken: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.url, accessToken, cmd, this.debug);
      })
      .then((contextResponse: ContextInfo): Promise<CustomAction[]> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(contextResponse));
          cmd.log('');

          cmd.log(`Attempt to get custom actions list with scope: ${args.options.scope}`);
          cmd.log('');
        }

        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.getCustomActions(args.options, siteAccessToken, cmd);
        }
        return this.searchAllScopes(args.options, siteAccessToken, cmd);
      })
      .then((customActions: CustomAction[]): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(customActions);
          cmd.log('');
        }

        if (customActions.length === 0) {
          if (this.verbose) {
            cmd.log(`Custom actions not found`);
          }
        }
        else {
          if (args.options.output === 'json') {
            cmd.log(customActions);
          }
          else {
            cmd.log(customActions.map(a => {
              return {
                Name: a.Name,
                Location: a.Location,
                Scope: this.humanizeScope(a.Scope),
                Id: a.Id
              };
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getCustomActions(options: Options, accessToken: string, cmd: CommandInstance): Promise<CustomAction[]> {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (this.debug) {
      cmd.log('Executing web request...');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<CustomAction[]>((resolve: (list: CustomAction[]) => void, reject: (error: any) => void): void => {
      request.get(requestOptions)
        .then((response: { value: CustomAction[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
  }

  /**
   * Two REST GET requests with `web` and `site` scope are sent.
   * The results are combined in one array.
   */
  private searchAllScopes(options: Options, accessToken: string, cmd: CommandInstance): Promise<CustomAction[]> {
    return new Promise<CustomAction[]>((resolve: (list: CustomAction[]) => void, reject: (error: any) => void): void => {
      options.scope = "Web";
      let webCustomActions: CustomAction[] = [];

      this
        .getCustomActions(options, accessToken, cmd)
        .then((customActions: CustomAction[]): Promise<CustomAction[]> => {
          if (this.debug) {
            cmd.log('getCustomActions with scope of web. Result...');
            cmd.log(JSON.stringify(customActions));
            cmd.log('');
          }

          webCustomActions = customActions;

          options.scope = "Site";

          return this.getCustomActions(options, accessToken, cmd);
        })
        .then((siteCustomActions: CustomAction[]): void => {
          if (this.debug) {
            cmd.log('getCustomActions with scope of site. Result...');
            cmd.log(JSON.stringify(siteCustomActions));
            cmd.log('');
          }

          resolve(siteCustomActions.concat(webCustomActions));
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'Url of the site (collection) to retrieve the custom action from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      if (SpoCommand.isValidSharePointUrl(args.options.url) !== true) {
        return 'Missing required option url';
      }

      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CUSTOMACTION_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
        using the ${chalk.blue(commands.CONNECT)} command.
                      
  Remarks:

    To retrieve list of custom actions, you have to first connect to a SharePoint Online site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    When using the text output type (default), the command lists only the values of the ${chalk.grey('Name')},
    ${chalk.grey('Location')}, ${chalk.grey('Scope')} and ${chalk.grey('Id')} properties of the custom action. When setting the output
    type to JSON, all available properties are included in the command output.

  Examples:
  
    Return details about all user custom actions located
    in site or site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_LIST} -u https://contoso.sharepoint.com/sites/test

    Return details about all user custom actions located
    in site or site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_LIST} --url https://contoso.sharepoint.com/sites/test

    Return details about all user custom actions located 
    in site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_LIST} -u https://contoso.sharepoint.com/sites/test -s Site

    Return details about all user custom actions located 
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_LIST} --url https://contoso.sharepoint.com/sites/test --scope Web

  More information:

    UserCustomAction REST API resources:
      https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction
      `);
  }
}

module.exports = new SpoCustomActionListCommand();