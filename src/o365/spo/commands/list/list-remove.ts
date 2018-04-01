import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
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
  webUrl: string;
  id?: string;
  title?: string;
  confirm?: boolean;
}

class SpoListRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    const removeList: () => void = (): void => {
      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          siteAccessToken = accessToken;

          if (this.verbose) {
            cmd.log(`Removing list in site at ${args.options.webUrl}...`);
          }

          let requestUrl: string = '';

          if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.id)}')`;
          }
          else {
            requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.title as string)}')`;
          }

          const requestOptions: any = {
            url: requestUrl,
            method: 'POST',
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeList();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the list ${args.options.id || args.options.title} from site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeList();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to remove is located'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the list to remove. Specify either id or title but not both'
      },
      {
        option: '-t, --title [title]',
        description: 'Title of the list to remove. Specify either id or title but not both'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the list'
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

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.id && args.options.title) {
        return 'Specify id or title, but not both';
      }

      if (!args.options.id && !args.options.title) {
        return 'Specify id or title';
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
  
    To remove a list, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Remove the list with ID ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf

    Remove the list with title ${chalk.grey('List 1')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
      ${chalk.grey(config.delimiter)} ${commands.LIST_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --title 'List 1'
      `);
  }
}

module.exports = new SpoListRemoveCommand();