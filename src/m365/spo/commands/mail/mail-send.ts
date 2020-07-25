import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoMailSendCommand();
