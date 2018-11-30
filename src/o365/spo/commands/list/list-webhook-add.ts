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
const expirationDateTimeMaxDays = 180;
const maxExpirationDateTime: Date = new Date();
// 180 days from now is the maximum expiration date for a webhook
maxExpirationDateTime.setDate(maxExpirationDateTime.getDate() + expirationDateTimeMaxDays);

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  notificationUrl: string;
  expirationDateTime?: string;
  clientState?: string;
}

class SpoListWebhookAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_ADD;
  }

  public get description(): string {
    return 'Adds a new webhook to the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.expirationDateTime = (!(!args.options.expirationDateTime)).toString();
    telemetryProps.clientState = (!(!args.options.clientState)).toString();
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
          cmd.log(`Retrieved access token ${accessToken}. Adding webhook to the specified list...`);
        }

        if (this.verbose) {
          cmd.log(`Adding webhook to list ${args.options.listId ? encodeURIComponent(args.options.listId) : encodeURIComponent(args.options.listTitle as string)} located at site ${args.options.webUrl}...`);
        }

        let requestUrl: string = '';

        if (args.options.listId) {
          requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions')`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions')`;
        }

        const requestBody: any = {};
        requestBody.resource = args.options.listId ? args.options.listId : args.options.listTitle;
        requestBody.notificationUrl = args.options.notificationUrl;
        // If no expiration date has been provided we will default to the
        // maximum expiration date of 180 days from now 
        requestBody.expirationDateTime = args.options.expirationDateTime
          ? new Date(args.options.expirationDateTime).toISOString()
          : maxExpirationDateTime.toISOString();
        if (args.options.clientState) {
          requestBody.clientState = args.options.clientState;
        }

        const requestOptions: any = {
          url: requestUrl,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: requestBody,
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to add the webhook to is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list to which the webhook which should be added. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list to which the webhook which should be added. Specify either listId or listTitle but not both'
      },
      {
        option: '-n, --notificationUrl <notificationUrl>',
        description: 'The notification url'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]',
        description: 'The expiration date. Will be set to max (6 months from today) if not provided.'
      },
      {
        option: '-c, --clientState [clientState]',
        description: 'A client state information that will be passed through notifications.'
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

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      if (!args.options.notificationUrl) {
        return 'Required parameter notificationUrl missing';
      }

      const parsedDateTime = Date.parse(args.options.expirationDateTime as string)
      if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
        return `Provide the date in one of the following formats:
      'YYYY-MM-DD'
      'YYYY-MM-DDThh:mm'
      'YYYY-MM-DDThh:mmZ'
      'YYYY-MM-DDThh:mm±hh:mm'`;
      }

      if (parsedDateTime < Date.now() || new Date(parsedDateTime) >= maxExpirationDateTime) {
        return `Provide an expiration date which is a date time in the future and within 6 months from now`;
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
  
    To add a webhook, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Add a web hook to the list ${chalk.grey('Documents')} located in site 
    ${chalk.grey('https://contoso.sharepoint.com/sites/ninja')} with the notification url 
    ${chalk.grey('https://contoso-functions.azurewebsites.net/webhook')} and the default expiration date
    ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_ADD} --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-functions.azurewebsites.net/webhook

    Add a web hook to the list ${chalk.grey('Documents')} located in site 
    ${chalk.grey('https://contoso.sharepoint.com/sites/ninja')} with the notification url 
    ${chalk.grey('https://contoso-functions.azurewebsites.net/webhook')} and an expiration date of ${chalk.grey('January 21st, 2019')}
    ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_ADD} --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-functions.azurewebsites.net/webhook --expirationDateTime 2019-01-21
    
    Add a web hook to the list ${chalk.grey('Documents')} located in site 
    ${chalk.grey('https://contoso.sharepoint.com/sites/ninja')} with the notification url 
    ${chalk.grey('https://contoso-functions.azurewebsites.net/webhook')}, a very specific expiration date
    of ${chalk.grey('6:15 PM on March 2nd, 2019')} and a client state
    ${chalk.grey(config.delimiter)} ${commands.LIST_WEBHOOK_ADD} --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-functions.azurewebsites.net/webhook --expirationDateTime '2019-03-02T18:15' --clientState "Hello State!"
      `);
  }
}

module.exports = new SpoListWebhookAddCommand();