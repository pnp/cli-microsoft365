import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The userId value must be a valid GUID.'
    }).optional().alias('u'),
  userName: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The userName value must be a valid user principal name (UPN).'
    }).optional().alias('n'),
  email: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The email value must be a valid email.'
    }).optional(),
  meetingId: z.string().alias('m')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingTranscriptListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_TRANSCRIPT_LIST;
  }

  public get description(): string {
    return 'Lists all transcripts for a given meeting';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'createdDateTime'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType {
    return schema.refine(options => [options.userId, options.userName, options.email].filter(value => value !== undefined).length <= 1, {
      message: 'Specify either userId, userName or email, but not multiple.',
      params: {
        customCode: 'optionSet',
        options: ['userId', 'userName', 'email']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving transcript list for the given meeting...`);
      }

      let requestUrl: string = `${this.resource}/beta/`;
      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName && !args.options.email) {
          throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting transcripts list using app only permissions`;
        }

        requestUrl += 'users/';
        if (args.options.userId) {
          requestUrl += args.options.userId;
        }
        else if (args.options.userName) {
          requestUrl += args.options.userName;
        }
        else if (args.options.email) {
          if (this.verbose) {
            await logger.logToStderr(`Getting user ID for user with email '${args.options.email}'.`);
          }
          const userId: string = await entraUser.getUserIdByEmail(args.options.email!);
          requestUrl += userId;
        }
      }
      else {
        if (args.options.userId || args.options.userName || args.options.email) {
          throw `The options 'userId', 'userName' and 'email' cannot be used while retrieving meeting transcripts using delegated permissions`;
        }

        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings/${args.options.meetingId}/transcripts`;
      const res = await odata.getAllItems<any>(requestUrl);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMeetingTranscriptListCommand();