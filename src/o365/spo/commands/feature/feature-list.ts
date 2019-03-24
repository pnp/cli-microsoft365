import auth from '../../SpoAuth';
import { Auth } from '../../../../Auth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { Feature } from './Feature';

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
    return 'Lists Features activated in the specified site or site collection';
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
      .then((accessToken: string): Promise<{ value: Feature[]; }> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving Features activated in ${scope}...`);
        }

        const requestOptions: any = {
          url: `${args.options.url}/_api/${scope}/Features?$select=DisplayName,DefinitionId`,
          headers: {
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((features: { value: Feature[] }): void => {
        if (features.value && features.value.length > 0) {
          cmd.log(features.value);
        }
        else {
          if (this.verbose) {
            cmd.log('No activated Features found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site (collection) to retrieve the activated Features from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the Features to retrieve. Allowed values Site|Web. Default Web',
        autocomplete: ['Site', 'Web']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web') {
          return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
        }
      }

      return SpoCommand.isValidSharePointUrl(args.options.url);
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FEATURE_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
                      
  Remarks:
  
    To retrieve list of activated Features, you have to first log in to
    a SharePoint Online site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
  
  Examples:
  
    Return details about Features activated in the specified site collection
      ${chalk.grey(config.delimiter)} ${commands.FEATURE_LIST} --url https://contoso.sharepoint.com/sites/test --scope Site

    Return details about Features activated in the specified site
      ${chalk.grey(config.delimiter)} ${commands.FEATURE_LIST} --url https://contoso.sharepoint.com/sites/test --scope Web
      `);
  }
}

module.exports = new SpoFeatureListCommand();