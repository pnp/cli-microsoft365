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
import { Feature } from './feature';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  scope?: string;
}

class SpoFeatureListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.FEATURE_LIST}`;
  }

  public get description(): string {
    return 'Lists features for site or site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'Web';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope : 'Web';
    const resource: string = Auth.getResourceFromUrl(args.options.url);
    let siteAccessToken: string = '';

    auth.getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.url, accessToken, cmd, this.debug);
      })
      .then((contextResponse: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(contextResponse));
          cmd.log('');

          cmd.log(`Attempt to get feature list with scope: ${scope}`);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${args.options.url}/_api/${scope}/Features?$select=DisplayName,DefinitionId`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        return request.get(requestOptions);
      })
      .then((features: { value: Feature[] }): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(features);
          cmd.log('');
        }

        if (features.value && features.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(features.value);
          }
          else {
            cmd.log(features.value.map(f => {
              return {
                DefinitionId: f.DefinitionId,
                DisplayName: f.DisplayName
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No features found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'Url of the site (collection) to retrieve the feature from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the feature. Allowed values Site|Web. Default Web',
        autocomplete: ['Site', 'Web']
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
          args.options.scope !== 'Web') {
          return `${args.options.scope} is not a valid feature scope. Allowed values are Site|Web`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FEATURE_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
        using the ${chalk.blue(commands.LOGIN)} command.
                      
  Remarks:
  
    To retrieve list of features, you have to first log in to a SharePoint Online site using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
    When using the text output type (default), the command lists only the values of the ${chalk.grey('DefinitionId')},
    and ${chalk.grey('DisplayName')} properties of the feature. When setting the output
    type to JSON, all available properties are included in the command output.
  
  Examples:
  
    Return details about all features located 
    in site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.FEATURE_LIST} -u https://contoso.sharepoint.com/sites/test -s Site

    Return details about all features located 
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.FEATURE_LIST} --url https://contoso.sharepoint.com/sites/test --scope Web

  More information:

    Feature REST API resources:
      https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj247054(v=office.15)#rest-resource-endpoint
      `);
  }
}

module.exports = new SpoFeatureListCommand();