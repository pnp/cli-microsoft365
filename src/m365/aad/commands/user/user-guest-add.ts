import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  emailAddress: string;
  displayName?: string;
  inviteRedirectUrl?: string;
  welcomeMessage?: string;
  messageLanguage?: string;
  ccRecipients?: string;
  sendInvitationMessage?: boolean;
}

class AadUserGuestAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GUEST_ADD;
  }

  public get description(): string {
    return 'Invite an external user to the organization';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        inviteRedirectUrl: typeof args.options.inviteRedirectUrl !== 'undefined',
        welcomeMessage: typeof args.options.welcomeMessage !== 'undefined',
        messageLanguage: typeof args.options.messageLanguage !== 'undefined',
        ccRecipients: typeof args.options.ccRecipients !== 'undefined',
        sendInvitationMessage: !!args.options.sendInvitationMessage
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--emailAddress <emailAddress>'
      },
      {
        option: '--displayName [displayName]'
      },
      {
        option: '--inviteRedirectUrl [inviteRedirectUrl]'
      },
      {
        option: '--welcomeMessage [welcomeMessage]'
      },
      {
        option: '--messageLanguage [messageLanguage]'
      },
      {
        option: '--ccRecipients [ccRecipients]'
      },
      {
        option: '--sendInvitationMessage'
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'inviteRedeemUrl', 'invitedUserDisplayName', 'invitedUserEmailAddress', 'invitedUserType', 'resetRedemption', 'sendInvitationMessage', 'status'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/invitations`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          invitedUserEmailAddress: args.options.emailAddress,
          inviteRedirectUrl: args.options.inviteRedirectUrl || 'https://myapplications.microsoft.com',
          invitedUserDisplayName: args.options.displayName,
          sendInvitationMessage: args.options.sendInvitationMessage,
          invitedUserMessageInfo: {
            customizedMessageBody: args.options.welcomeMessage,
            messageLanguage: args.options.messageLanguage || 'en-US',
            ccRecipients: args.options.ccRecipients ? this.mapEmailsToRecipients([args.options.ccRecipients]) : undefined
          }
        }
      };

      const result = await request.post<any>(requestOptions);
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapEmailsToRecipients(emails: string[]): { emailAddress: { address: string }; }[] {
    return emails.map(mail => ({
      emailAddress: {
        address: mail.trim()
      }
    }));
  }
}

module.exports = new AadUserGuestAddCommand();