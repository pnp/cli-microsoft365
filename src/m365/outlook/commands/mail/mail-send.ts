import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
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
  sender?: string;
  mailbox?: string;
  bodyContents?: string;
  bodyContentType?: string;
  saveToSentItems?: string;
}

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an e-mail';
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
        bodyContents: typeof args.options.bodyContents !== 'undefined',
        bodyContentType: args.options.bodyContentType,
        saveToSentItems: args.options.saveToSentItems,
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

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    const bodyContents: string = args.options.bodyContents as string;

    const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAuth === true && !args.options.sender) {
      return cb(new CommandError(`Specify a upn or user id in the 'sender' option when using app only authentication.`));
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

    if (args.options.mailbox) {
      requestOptions.data.message.from = {
        emailAddress: {
          address: args.options.mailbox
        }
      };
    }

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new OutlookMailSendCommand();