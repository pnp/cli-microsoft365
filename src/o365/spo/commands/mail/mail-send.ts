import auth from '../../SpoAuth';
import { ContextInfo } from '../../spo';
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

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
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
    return 'Send an email from SharePoint';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.to = typeof args.options.to !== 'undefined';
    telemetryProps.subject = typeof args.options.subject !== 'undefined';
    telemetryProps.body = typeof args.options.body !== 'undefined';
    telemetryProps.from = typeof args.options.from !== 'undefined';
    telemetryProps.cc = typeof args.options.cc !== 'undefined';
    telemetryProps.bcc = typeof args.options.bcc !== 'undefined';
    telemetryProps.additionalHeaders = typeof args.options.additionalHeaders !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.debug) {
      cmd.log(`Retrieving access token for the site collection ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const params: any = {};
        params.properties = {};
        params.properties.__metadata = { "type": "SP.Utilities.EmailProperties" };
        params.properties.Body = args.options.body;
        params.properties.Subject = args.options.subject

        params.properties.To = { results: args.options.to.replace(/\s+/g, '').split(',') };
        if (this.verbose) {
          cmd.log(`List of recipients: ${args.options.to}...`);
        }

        if (args.options.from && args.options.from.length > 0) {
          params.properties.From = args.options.from;
          if (this.verbose) {
            cmd.log(`Mail will send from: ${args.options.from}...`);
          }
        }

        if (args.options.cc && args.options.cc.length > 0) {
          params.properties.CC = { results: args.options.cc.replace(/\s+/g, '').split(',') };
          if (this.verbose) {
            cmd.log(`List of addresses to which a carbon copy: ${args.options.cc}...`);
          }
        }

        if (args.options.bcc && args.options.bcc.length > 0) {
          params.properties.BCC = { results: args.options.bcc.replace(/\s+/g, '').split(',') };
          if (this.verbose) {
            cmd.log(`List of addresses that receive a copy of the mail but are not listed as recipients: ${args.options.bcc}...`);
          }
        }

        if (args.options.additionalHeaders) {
          params.properties.AdditionalHeaders = args.options.additionalHeaders;
          if (this.verbose) {
            cmd.log(`Additional headers informaitons: ${args.options.additionalHeaders}...`);
          }
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/SP.Utilities.Utility.SendEmail`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue,
            'Accept': 'application/json;odata=verbose',
            'content-type': 'application/json;odata=verbose'
          }),
          body: JSON.stringify(params)
        };

        if (this.debug) {
          cmd.log('Executing the request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--to <to>',
        description: 'Recipient\'s email address (separate recipients by comma)'
      },
      {
        option: '--subject <subject>',
        description: 'Subject of the email'
      },
      {
        option: '--body <body>',
        description: 'Content of the email'
      },
      {
        option: '--from [from]',
        description: 'Sender\'s email address'
      },
      {
        option: '--cc [cc]',
        description: 'Email addresses to which a carbon copy (CC) of the email is sent (separate addresses by comma)'
      },
      {
        option: '--bcc [bcc]',
        description: 'Email addresses that receive a copy of the mail but are not listed as recipients of the message (separate addresses by comma)'
      },
      {
        option: '--additionalHeaders [additionalHeaders]',
        description: 'Add additional headers informations'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.to) {
        return 'Specify at least one recipient, is required';
      }

      if (!args.options.subject) {
        return 'Specify an email subject, is required';
      }

      if (!args.options.body) {
        return 'Specify the content of the email, is required';
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
  
  To send an email, you have to first log in to a SharePoint Online site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

  All recipients (internal and external) have to have access to the target SharePoint site.
        
  Examples:
  
    Send an email to ${chalk.grey('user@contoso.com')} 
      ${chalk.grey(config.delimiter)} ${commands.MAIL_SEND} --to 'user@contoso.com' --subject 'Email send via Office365-CLI' --body '<h1>Office365-CLI</h1>Email send via <b>cmdlet</b>.'
    
    Send an email to multiples addresses
      ${chalk.grey(config.delimiter)} ${commands.MAIL_SEND} --to 'user1@contoso.com,user2@contoso.com' --subject 'Email send via Office365-CLI' --body '<h1>Office365-CLI</h1>Email send via <b>cmdlet</b>.' --cc 'user3@contoso.com' --bcc 'user4@contoso.com'
      `);
  }
}

module.exports = new SpoMailSendCommand();