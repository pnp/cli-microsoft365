import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends e-mail on behalf of the current user';
  }

  public alias(): string[] | undefined {
    return [commands.SENDMAIL];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.bodyContents = typeof args.options.bodyContents !== 'undefined';
    telemetryProps.bodyContentsFilePath = typeof args.options.bodyContentsFilePath !== 'undefined';
    telemetryProps.bodyContentType = args.options.bodyContentType;
    telemetryProps.saveToSentItems = args.options.saveToSentItems;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
      responseType: 'json',
      data: {
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
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --subject <subject>'
      },
      {
        option: '-t, --to <to>'
      },
      {
        option: '--bodyContents [bodyContents]'
      },
      {
        option: '--bodyContentsFilePath [bodyContentsFilePath]'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['Text', 'HTML']
      },
      {
        option: '--saveToSentItems [saveToSentItems]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new OutlookMailSendCommand();
