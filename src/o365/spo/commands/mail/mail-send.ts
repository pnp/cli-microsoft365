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
  to: string;
  subject: string;
  body: string;
  from?: string;
  cc?: string;
  bcc?: string;
  additionalHeaders?: string;
}

class SpoMailSendCommand extends SpoCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an e-mail from SharePoint';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.from = typeof args.options.from !== 'undefined';
    telemetryProps.cc = typeof args.options.cc !== 'undefined';
    telemetryProps.bcc = typeof args.options.bcc !== 'undefined';
    telemetryProps.additionalHeaders = typeof args.options.additionalHeaders !== 'undefined';
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
      .then((accessToken: string): request.RequestPromise | Promise<void> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${siteAccessToken}.`);
        }

        const params: any = {
          properties: {
            __metadata: { "type": "SP.Utilities.EmailProperties" },
            Body: args.options.body,
            Subject: args.options.subject,
            To: { results: args.options.to.replace(/\s+/g, '').split(',') }
          }
        };

        if (args.options.from && args.options.from.length > 0) {
          params.properties.From = args.options.from;
        }

        if (args.options.cc && args.options.cc.length > 0) {
          params.properties.CC = { results: args.options.cc.replace(/\s+/g, '').split(',') };
        }

        if (args.options.bcc && args.options.bcc.length > 0) {
          params.properties.BCC = { results: args.options.bcc.replace(/\s+/g, '').split(',') };
        }

        if (args.options.additionalHeaders) {
          const h = JSON.parse(args.options.additionalHeaders);
          params.properties.AdditionalHeaders = {
            __metadata: { "type": "Collection(SP.KeyValue)" },
            results: Object.keys(h).map(key => {
              return {
                __metadata: {
                  type: 'SP.KeyValue'
                },
                Key: key,
                Value: h[key],
                ValueType: 'Edm.String'
              }
            })
          };
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/SP.Utilities.Utility.SendEmail`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'content-type': 'application/json;odata=verbose'
          }),
          json: true,
          body: params
        };

        if (this.debug) {
          cmd.log('Executing the request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((rawRes: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(rawRes);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site from which the e-mail will be sent'
      },
      {
        option: '--to <to>',
        description: 'Comma-separated list of recipients\' e-mail addresses'
      },
      {
        option: '--subject <subject>',
        description: 'Subject of the e-mail'
      },
      {
        option: '--body <body>',
        description: 'Content of the e-mail'
      },
      {
        option: '--from [from]',
        description: 'Sender\'s e-mail address'
      },
      {
        option: '--cc [cc]',
        description: 'Comma-separated list of CC recipients'
      },
      {
        option: '--bcc [bcc]',
        description: 'Comma-separated list of BCC recipients'
      },
      {
        option: '--additionalHeaders [additionalHeaders]',
        description: 'JSON string with additional headers'
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

      if (!args.options.to) {
        return 'Required parameter to missing';
      }

      if (!args.options.subject) {
        return 'Required parameter subject missing';
      }

      if (!args.options.body) {
        return 'Required parameter body missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site using
    the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To send an email, you have to first log in to a SharePoint Online site using
    the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    All recipients (internal and external) have to have access to the target
    SharePoint site.
        
  Examples:
  
    Send an e-mail to ${chalk.grey('user@contoso.com')} 
      ${chalk.grey(config.delimiter)} ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to 'user@contoso.com' --subject 'Email sent via Office 365 CLI' --body '<h1>Office 365 CLI</h1>Email sent via <b>command</b>.'
    
    Send an e-mail to multiples addresses
      ${chalk.grey(config.delimiter)} ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to 'user1@contoso.com,user2@contoso.com' --subject 'Email sent via Office 365 CLI' --body '<h1>Office 365 CLI</h1>Email sent via <b>command</b>.' --cc 'user3@contoso.com' --bcc 'user4@contoso.com'
    
    Send an e-mail to ${chalk.grey('user@contoso.com')} with additional headers
      ${chalk.grey(config.delimiter)} ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to 'user@contoso.com' --subject 'Email sent via Office 365 CLI' --body '<h1>Office 365 CLI</h1>Email sent via <b>command</b>.' --additionalHeaders "'{\"X-MC-Tags\":\"Office 365 CLI\"}'"
      `);
  }
}

module.exports = new SpoMailSendCommand();