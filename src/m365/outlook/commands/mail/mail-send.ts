import * as fs from 'fs';
import * as path from 'path';
import { AxiosRequestConfig } from 'axios';
import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import { formatting } from '../../../../utils/formatting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  subject: string;
  to: string;
  cc?: string;
  bcc?: string;
  sender?: string;
  mailbox?: string;
  bodyContents: string;
  bodyContentType?: string;
  importance?: string;
  attachment?: string | string[];
  saveToSentItems?: boolean;
}

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an email';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        cc: typeof args.options.cc !== 'undefined',
        bcc: typeof args.options.bcc !== 'undefined',
        bodyContentType: args.options.bodyContentType,
        saveToSentItems: args.options.saveToSentItems,
        importance: args.options.importance,
        mailbox: typeof args.options.mailbox !== 'undefined',
        sender: typeof args.options.sender !== 'undefined',
        attachment: typeof args.options.attachment !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --subject <subject>'
      },
      {
        option: '-t, --to <to>'
      },
      {
        option: '--cc [cc]'
      },
      {
        option: '--bcc [bcc]'
      },
      {
        option: '--sender [sender]'
      },
      {
        option: '-m, --mailbox [mailbox]'
      },
      {
        option: '--bodyContents <bodyContents>'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['Text', 'HTML']
      },
      {
        option: '--importance [importance]',
        autocomplete: ['low', 'normal', 'high']
      },
      {
        option: '--attachment [attachment]'
      },
      {
        option: '--saveToSentItems [saveToSentItems]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.bodyContentType &&
          args.options.bodyContentType !== 'Text' &&
          args.options.bodyContentType !== 'HTML') {
          return `${args.options.bodyContentType} is not a valid value for the bodyContentType option. Allowed values are Text|HTML`;
        }

        if (args.options.saveToSentItems && !validation.isValidBoolean(args.options.saveToSentItems as any)) {
          return `${args.options.saveToSentItems} is not a valid value for the saveToSentItems option. Allowed values are true|false`;
        }

        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance) === -1) {
          return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
        }

        if (args.options.attachment) {
          const attachments: string[] = typeof args.options.attachment === 'string' ? [args.options.attachment] : args.options.attachment;

          for (const attachment of attachments) {
            if (!fs.existsSync(attachment)) {
              return `File with path '${attachment}' was not found.`;
            }

            if (!fs.lstatSync(attachment).isFile()) {
              return `'${attachment}' is not a file.`;
            }
          }

          const requestBody = this.getRequestBody(args.options);
          // The max body size of the request is 4 194 304 chars before getting a 413 response
          if (JSON.stringify(requestBody).length > 4_194_304) {
            return 'Exceeded the max total size of attachments which is 3MB.';
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[this.resource].accessToken);
      if (isAppOnlyAuth === true && !args.options.sender) {
        throw `Specify a upn or user id in the 'sender' option when using app only authentication.`;
      }

      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/${args.options.sender ? 'users/' + formatting.encodeQueryParameter(args.options.sender) : 'me'}/sendMail`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: this.getRequestBody(args.options)
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapEmailAddressToRecipient(email: string | undefined): { emailAddress: { address: string }; } | undefined {
    if (!email) {
      return undefined;
    }

    return {
      emailAddress: {
        address: email.trim()
      }
    };
  }

  private getRequestBody(options: Options): { message: any, saveToSentItems?: boolean } {
    const attachments = typeof options.attachment === 'string' ? [options.attachment] : options.attachment;
    const attachmentContents = attachments?.map(a => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: path.basename(a),
      contentBytes: fs.readFileSync(a, { encoding: 'base64' })
    }));

    return ({
      message: {
        subject: options.subject,
        body: {
          contentType: options.bodyContentType || 'Text',
          content: options.bodyContents
        },
        from: this.mapEmailAddressToRecipient(options.mailbox),
        toRecipients: options.to.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
        ccRecipients: options.cc?.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
        bccRecipients: options.bcc?.split(',').map(mail => this.mapEmailAddressToRecipient(mail)),
        importance: options.importance,
        attachments: attachmentContents
      },
      saveToSentItems: options.saveToSentItems
    });
  }
}

module.exports = new OutlookMailSendCommand();