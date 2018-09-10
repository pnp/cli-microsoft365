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
}

class SpoListWebhookListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_LIST;
  }

  public get description(): string {
    return 'Lists all webhooks for the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about webhooks...`);
        }

        if (this.verbose) {
          const list: string = args.options.id ? encodeURIComponent(args.options.id as string) : encodeURIComponent(args.options.title as string);
          cmd.log(`Retrieving webhook information for list ${list} in site at ${args.options.webUrl}...`);
        }

        let requestUrl: string = '';

        if (args.options.id) {
          requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.id)}')/Subscriptions`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.title as string)}')/Subscriptions`;
        }

        const requestOptions: any = {
          url: requestUrl,
          method: 'GET',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
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
      .then((res: { value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(res.value);
          }
          else {
            cmd.log(res.value.map(e => {
              return {
                id: e.id,
                clientState: e.clientState || '',
                expirationDateTime: e.expirationDateTime,
                resource: e.resource
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No webhooks found');
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to retrieve webhooks for is located'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the list to retrieve all webhooks for. Specify either id or title but not both'
      },
      {
        option: '-t, --title [title]',
        description: 'Title of the list to retrieve all webhooks for. Specify either id or title but not both'
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

      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (args.options.id && args.options.title) {
        return 'Specify id or title, but not both';
      }

      if (!args.options.id && !args.options.title) {
        return 'Specify id or title, one is required';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to SharePoint,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To list all webhooks for a list, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    List all webhooks for a list with ID ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf

    List all webhooks for a list with title ${chalk.grey('Documents')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x --title Documents
      `);
  }
}

module.exports = new SpoListWebhookListCommand();