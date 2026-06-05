import fs from 'fs';
import path from 'path';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  subject: z.string().alias('s'),
  to: z.string().alias('t'),
  cc: z.string().optional(),
  bcc: z.string().optional(),
  sender: z.string().optional(),
  mailbox: z.string().optional().alias('m'),
  bodyContents: z.string(),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  attachment: z.union([z.string(), z.string().array()]).optional(),
  saveToSentItems: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMailSendCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEND;
  }

  public get description(): string {
    return 'Sends an email';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(options => {
        if (options.attachment) {
          const attachments: string[] = typeof options.attachment === 'string' ? [options.attachment] : options.attachment;

          for (const attachment of attachments) {
            if (!fs.existsSync(attachment)) {
              return false;
            }
          }
        }

        return true;
      }, {
        error: 'One or more attachment files were not found.'
      })
      .refine(options => {
        if (options.attachment) {
          const attachments: string[] = typeof options.attachment === 'string' ? [options.attachment] : options.attachment;

          for (const attachment of attachments) {
            if (fs.existsSync(attachment) && !fs.lstatSync(attachment).isFile()) {
              return false;
            }
          }
        }

        return true;
      }, {
        error: 'One or more attachments is not a file.'
      })
      .refine(options => {
        if (options.attachment) {
          const requestBody = this.getRequestBody(options);
          // The max body size of the request is 4 194 304 chars before getting a 413 response
          if (JSON.stringify(requestBody).length > 4_194_304) {
            return false;
          }
        }

        return true;
      }, {
        error: 'Exceeded the max total size of attachments which is 3MB.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
      if (isAppOnlyAccessToken === true && !args.options.sender) {
        throw `Specify a upn or user id in the 'sender' option when using app only authentication.`;
      }

      const requestOptions: CliRequestOptions = {
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

export default new OutlookMailSendCommand();