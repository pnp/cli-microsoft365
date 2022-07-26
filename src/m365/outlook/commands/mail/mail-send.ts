import * as fs from 'fs';
import * as path from 'path';
import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
  bodyContents: string;
  bodyContentType?: string;
  saveToSentItems?: string;
  attachment?: string | string[];
}

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an email on behalf of the current user';
  }

  public alias(): string[] | undefined {
    return [commands.SENDMAIL];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.bodyContents = typeof args.options.bodyContents !== 'undefined';
    telemetryProps.bodyContentType = args.options.bodyContentType;
    telemetryProps.saveToSentItems = typeof args.options.saveToSentItems !== 'undefined';
    telemetryProps.attachment = typeof args.options.attachment !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const attachmentPaths: string[] | undefined = typeof args.options.attachment === 'string' ? [args.options.attachment] : args.options.attachment;

    const attachments = attachmentPaths?.map(a => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: path.basename(a),
      contentBytes: fs.readFileSync(a, { encoding: 'base64' })
    }));

    const requestOptions: AxiosRequestConfig = {
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
            content: args.options.bodyContents
          },
          toRecipients: args.options.to.split(',').map(e => ({
            emailAddress: {
              address: e.trim()
            }
          })),
          attachments: attachments
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
        option: '--bodyContents <bodyContents>'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['Text', 'HTML']
      },
      {
        option: '--saveToSentItems [saveToSentItems]'
      },
      {
        option: '--attachment [attachment]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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

    if (args.options.attachment) {
      let totalSize: number = 0;
      const attachmentPaths: string[] = typeof args.options.attachment === 'string' ? [args.options.attachment] : args.options.attachment;
      for (const attachment of attachmentPaths) {
        if (!fs.existsSync(attachment)) {
          return `File with path '${attachment}' was not found.`;
        }

        const fileInfo = fs.lstatSync(attachment);
        if (!fileInfo.isFile()) {
          return `'${attachment}' is not a file.`;
        }

        totalSize += fileInfo.size;
      }

      if (totalSize > 3_100_000) {
        return `Exceeded the max total size of attachments which is 3 MB.`;
      }
    }

    return true;
  }
}

module.exports = new OutlookMailSendCommand();
