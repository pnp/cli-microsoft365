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
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';
import { CustomAction } from './customaction';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  url: string;
  scope?: string;
}

class SpoCustomActionGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_GET}`;
  }

  public get description(): string {
    return 'Gets details for the specified custom action';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.url);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.url, siteAccessToken, cmd, this.debug);
      })
      .then((contextResponse: ContextInfo): request.RequestPromise | Promise<CustomAction> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(contextResponse));
          cmd.log('');

          cmd.log(`Attempt to get custom action with scope: ${args.options.scope}`);
          cmd.log('');
        }

        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          return this.getCustomAction(args.options, siteAccessToken, cmd);
        }

        return this.searchAllScopes(args.options, siteAccessToken, cmd);
      })
      .then((customAction: CustomAction): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(customAction));
          cmd.log('');
        }

        if (customAction["odata.null"] === true) {
          if (this.verbose) {
            cmd.log(`Custom action with id ${args.options.id} not found`);
          }
        }
        else {
          cmd.log({
            ClientSideComponentId: customAction.ClientSideComponentId,
            ClientSideComponentProperties: customAction.ClientSideComponentProperties,
            CommandUIExtension: customAction.CommandUIExtension,
            Description: customAction.Description,
            Group: customAction.Group,
            Id: customAction.Id,
            ImageUrl: customAction.ImageUrl,
            Location: customAction.Location,
            Name: customAction.Name,
            RegistrationId: customAction.RegistrationId,
            RegistrationType: customAction.RegistrationType,
            Rights: JSON.stringify(customAction.Rights),
            Scope: this.humanizeScope(customAction.Scope),
            ScriptBlock: customAction.ScriptBlock,
            ScriptSrc: customAction.ScriptSrc,
            Sequence: customAction.Sequence,
            Title: customAction.Title,
            Url: customAction.Url,
            VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getCustomAction(options: Options, siteAccessToken: string, cmd: CommandInstance): request.RequestPromise {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions('${encodeURIComponent(options.id)}')`,
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

    return request.get(requestOptions);
  }

  /**
   * Get request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private searchAllScopes(options: Options, siteAccessToken: string, cmd: CommandInstance): Promise<CustomAction> {
    return new Promise<CustomAction>((resolve: (customAction: CustomAction) => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .getCustomAction(options, siteAccessToken, cmd)
        .then((webResult: CustomAction): void => {
          if (this.debug) {
            cmd.log('getCustomAction with scope of web result...');
            cmd.log(JSON.stringify(webResult));
            cmd.log('');
          }

          if (webResult["odata.null"] !== true) {
            return resolve(webResult);
          }

          options.scope = "Site";
          this
            .getCustomAction(options, siteAccessToken, cmd)
            .then((siteResult: CustomAction): void => {
              if (this.debug) {
                cmd.log('getCustomAction with scope of site result...');
                cmd.log(JSON.stringify(siteResult));
                cmd.log('');
              }

              return resolve(siteResult);
            }, (err: any): void => {
              reject(err);
            });
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
        option: '-i, --id <id>',
        description: 'Id (Guid) of the custom action to retrieve'
      },
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
      if (Utils.isValidGuid(args.options.id) === false) {
        return `${args.options.id} is not valid. Custom action id (Guid) expected.`;
      }

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
    log(vorpal.find(commands.CUSTOMACTION_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
        using the ${chalk.blue(commands.CONNECT)} command.
                      
  Remarks:

    To retrieve custom action, you have to first connect to a SharePoint Online site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:
  
    Return details about the user custom action with ID ${chalk.grey('058140e3-0e37-44fc-a1d3-79c487d371a3')}
    located in site or site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test

    Return details about the user custom action with ID ${chalk.grey('058140e3-0e37-44fc-a1d3-79c487d371a3')}
    located in site or site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test

    Return details about the user custom action with ID ${chalk.grey('058140e3-0e37-44fc-a1d3-79c487d371a3')}
    located in site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test -s Site

    Return details about the user custom action with ID ${chalk.grey('058140e3-0e37-44fc-a1d3-79c487d371a3')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.CUSTOMACTION_GET} --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --scope Web

  More information:

    UserCustomAction REST API resources:
      https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction
      `);
  }
}

module.exports = new SpoCustomActionGetCommand();