import commands from '../../commands';
import * as path from 'path';
import * as fs from 'fs';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  subject: string;
  to: string;
  bodyContents?: string;
  bodyContentsFilePath?: string;
  bodyContentType?: string;
  saveToSentItems?: string;
}

class OutlookSendmailCommand extends GraphCommand {
  public get name(): string {
    return `${commands.OUTLOOK_MAIL_SEND}`;
  }

  public get description(): string {
    return 'Sends e-mail on behalf of the current user';
  }

  public alias(): string[] | undefined {
    return [commands.OUTLOOK_SENDMAIL];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.bodyContents = typeof args.options.bodyContents !== 'undefined';
    telemetryProps.bodyContentsFilePath = typeof args.options.bodyContentsFilePath !== 'undefined';
    telemetryProps.bodyContentType = args.options.bodyContentType;
    telemetryProps.saveToSentItems = args.options.saveToSentItems;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let bodyContents: string = args.options.bodyContents as string;
    if (args.options.bodyContentsFilePath) {
      bodyContents = fs.readFileSync(path.resolve(args.options.bodyContentsFilePath), 'utf-8');
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/sendMail`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      json: true,
      body: {
        message: {
          subject: args.options.subject,
          body: {
            contentType: args.options.bodyContentType || 'Text',
            content: bodyContents
          },
          toRecipients: args.options.to.split(',').map(e => {
            return {
              emailAddress: {
                address: e.trim()
              }
            };
          })
        },
        saveToSentItems: args.options.saveToSentItems
      }
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --subject <subject>',
        description: 'E-mail subject'
      },
      {
        option: '-t, --to <to>',
        description: 'Comma-separated list of e-mails to send the message to'
      },
      {
        option: '--bodyContents [bodyContents]',
        description: 'String containing the body of the e-mail to send'
      },
      {
        option: '--bodyContentsFilePath [bodyContentsFilePath]',
        description: 'Relative or absolute path to the file with e-mail body contents'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        description: 'Type of the body content. Available options: Text|HTML. Default Text',
        autocomplete: ['Text', 'HTML']
      },
      {
        option: '--saveToSentItems [saveToSentItems]',
        description: 'Save e-mail in the sent items folder. Default true'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.subject) {
        return 'Required option subject missing';
      }

      if (!args.options.to) {
        return 'Required option to missing';
      }

      if (!args.options.bodyContents && !args.options.bodyContentsFilePath) {
        return 'Specify either bodyContents or bodyContentsFilePath';
      }

      if (args.options.bodyContents && args.options.bodyContentsFilePath) {
        return 'Specify either bodyContents or bodyContentsFilePath but not both';
      }

      if (args.options.bodyContentsFilePath) {
        const fullPath: string = path.resolve(args.options.bodyContentsFilePath);

        if (!fs.existsSync(fullPath)) {
          return `File '${fullPath}' not found`;
        }

        if (fs.lstatSync(fullPath).isDirectory()) {
          return `Path '${fullPath}' points to a directory`;
        }
      }

      if (args.options.bodyContentType &&
        args.options.bodyContentType !== 'Text' &&
        args.options.bodyContentType !== 'HTML') {
        return `${args.options.bodyContents} is not a valid value for the bodyContents option. Allowed values are Text|HTML`;
      }

      if (args.options.saveToSentItems &&
        args.options.saveToSentItems !== 'true' &&
        args.options.saveToSentItems !== 'false') {
        return `${args.options.saveToSentItems} is not a valid value for the saveToSentItems option. Allowed values are true|false`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Send a text e-mail to the specified e-mail address
      ${this.name} --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site"

    Send an HTML e-mail to the specified e-mail addresses
      ${this.name} --to "chris@contoso.com,brian@contoso.com" --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the <a href='https://contoso.sharepoint.com/sites/marketing'>team site</a>" --bodyContentType HTML

    Send an HTML e-mail to the specified e-mail address loading e-mail contents
    from a file on disk
      ${this.name} --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContentsFilePath email.html --bodyContentType HTML

    Send a text e-mail to the specified e-mail address. Don't store the e-mail
    in sent items
      ${this.name} --to chris@contoso.com --subject "DG2000 Data Sheets" --bodyContents "The latest data sheets are in the team site" --saveToSentItems false
`);
  }
}

module.exports = new OutlookSendmailCommand();
