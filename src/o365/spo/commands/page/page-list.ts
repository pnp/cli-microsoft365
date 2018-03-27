import auth from '../../SpoAuth';
import { Auth } from '../../../../Auth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoPageListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_LIST}`;
  }

  public get description(): string {
    return 'Lists all modern pages in the given site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving client-side pages...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/lists/SitePages/rootfolder/files?$expand=ListItemAllFields/ClientSideApplicationId&$orderby=Name`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
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
      .then((res: { value: any[] }): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (res.value && res.value.length > 0) {
          const clientSidePages: any[] = res.value.filter(p => p.ListItemAllFields.ClientSideApplicationId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec');
          if (args.options.output === 'json') {
            cmd.log(clientSidePages);
          }
          else {
            cmd.log(clientSidePages.map(p => {
              return {
                Name: p.Name,
                Title: p.Title
              }
            }));
          }
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site from which to retrieve available pages'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Missing required option webUrl';
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To list all modern pages in the specific site, you have to first connect to
    a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:
  
    List all modern pages in the specific site
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a
`);
  }
}

module.exports = new SpoPageListCommand();