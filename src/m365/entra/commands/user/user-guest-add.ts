import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  emailAddress: z.string(),
  displayName: z.string().optional(),
  inviteRedirectUrl: z.string().optional(),
  welcomeMessage: z.string().optional(),
  messageLanguage: z.string().optional(),
  ccRecipients: z.string().optional(),
  sendInvitationMessage: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserGuestAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GUEST_ADD;
  }

  public get description(): string {
    return 'Invite an external user to the organization';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
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
      await logger.log(result);
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

export default new EntraUserGuestAddCommand();