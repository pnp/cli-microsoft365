import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli';
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
  cc?: string;
  bcc?: string;
  sender?: string;
  mailbox?: string;
  bodyContents: string;
  bodyContentType?: string;
  importance?: string;
  saveToSentItems?: string;
}

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an email';
  }

  public alias(): string[] | undefined {
    return [commands.SENDMAIL];
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
        sender: typeof args.options.sender !== 'undefined'
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
          return `${args.options.bodyContents} is not a valid value for the bodyContents option. Allowed values are Text|HTML`;
        }

        if (args.options.saveToSentItems &&
          args.options.saveToSentItems !== 'true' &&
          args.options.saveToSentItems !== 'false') {
          return `${args.options.saveToSentItems} is not a valid value for the saveToSentItems option. Allowed values are true|false`;
        }

        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance) === -1) {
          return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
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
  
      const requestOptions: any = {
        url: `${this.resource}/v1.0/${args.options.sender ? 'users/' + encodeURIComponent(args.options.sender) : 'me'}/sendMail`,
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
            toRecipients: this.mapEmailAddressesToRecipients(args.options.to.split(',')),
            ccRecipients: this.mapEmailAddressesToRecipients(args.options.cc?.split(',')),
            bccRecipients: this.mapEmailAddressesToRecipients(args.options.bcc?.split(',')),
            importance: args.options.importance
          },
          saveToSentItems: args.options.saveToSentItems
        }
      };
  
      if (args.options.mailbox) {
        requestOptions.data.message.from = {
          emailAddress: {
            address: args.options.mailbox
          }
        };
      }
  
      await request.post(requestOptions);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapEmailAddressesToRecipients(emailAddresses: string[] | undefined): { emailAddress: { address: string }; }[] | undefined {
    if (!emailAddresses) {
      return emailAddresses;
    }

    return emailAddresses.map(email => ({
      emailAddress: {
        address: email.trim()
      }
    }));
  }
}

module.exports = new OutlookMailSendCommand();