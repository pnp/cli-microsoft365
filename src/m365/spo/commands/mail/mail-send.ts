import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';

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
      headers: {
        'content-type': 'application/json;odata=verbose'
      },
      json: true,
      body: params
    };

    request
      .post(requestOptions)
      .then((): void => {
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
      `  Remarks:

    All recipients (internal and external) have to have access to the target
    SharePoint site.

  Examples:

    Send an e-mail to ${chalk.grey('user@contoso.com')}
      ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to "user@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>."

    Send an e-mail to multiples addresses
      ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to "user1@contoso.com,user2@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>." --cc "user3@contoso.com" --bcc "user4@contoso.com"

    Send an e-mail to ${chalk.grey('user@contoso.com')} with additional headers
      ${commands.MAIL_SEND} --webUrl https://contoso.sharepoint.com/sites/project-x --to "user@contoso.com" --subject "Email sent via CLI for Microsoft 365" --body "<h1>CLI for Microsoft 365</h1>Email sent via <b>command</b>." --additionalHeaders "'{\"X-MC-Tags\":\"CLI for Microsoft 365\"}'"
      `);
  }
}

module.exports = new SpoMailSendCommand();
