import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { MeetingTranscript } from '../../MeetingTranscript.js';
import fs from 'fs';
import path from 'path';

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
  meetingId: z.string().alias('m'),
  id: z.string().alias('i'),
  outputFile: z.string()
    .refine(value => fs.existsSync(path.dirname(value)), {
      message: 'Specified path where to save the file does not exist.'
    }).optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingTranscriptGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_TRANSCRIPT_GET;
  }

  public get description(): string {
    return 'Downloads a transcript for a given meeting';
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
        await logger.logToStderr(`Retrieving transcript for the given meeting...`);
      }

      let requestUrl: string = `${this.resource}/beta/`;
      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName && !args.options.email) {
          throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting transcript using app only permissions`;
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
          throw `The options 'userId', 'userName', and 'email' cannot be used while retrieving meeting transcript using delegated permissions`;
        }

        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings/${args.options.meetingId}/transcripts/${args.options.id}`;

      if (args.options.outputFile) {
        requestUrl += '/content?$format=text/vtt';
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: args.options.outputFile ? 'stream' : 'json'
      };

      const meetingTranscript = await request.get<MeetingTranscript>(requestOptions);

      if (meetingTranscript) {
        if (args.options.outputFile) {
          // Not possible to use async/await for this promise
          await new Promise<void>((resolve, reject) => {
            const writer = fs.createWriteStream(args.options.outputFile as string);
            (meetingTranscript as any).data.pipe(writer);

            writer.on('error', err => {
              reject(err);
            });

            writer.on('close', async () => {
              const filePath = args.options.outputFile as string;
              if (this.verbose) {
                await logger.logToStderr(`File saved to path ${filePath}`);
              }
              return resolve();
            });
          });
        }
        else {
          await logger.log(meetingTranscript);
        }
      }
      else {
        throw `The specified meeting transcript was not found`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMeetingTranscriptGetCommand();