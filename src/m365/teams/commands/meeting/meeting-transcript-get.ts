import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
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

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  email?: string;
  meetingId: string;
  id: string;
  outputFile?: string;
}

class TeamsMeetingTranscriptGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_TRANSCRIPT_GET;
  }

  public get description(): string {
    return 'Downloads a transcript for a given meeting';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        outputFile: typeof args.options.outputFile !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --userId [userId]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--email [email]'
      },
      {
        option: '-m, --meetingId <meetingId>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --outputFile [outputFile]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid Guid`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email`;
        }

        if (args.options.outputFile && !fs.existsSync(path.dirname(args.options.outputFile))) {
          return 'Specified path where to save the file does not exits';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName', 'email'],
      runsWhen: (args) => args.options.userId || args.options.userName || args.options.email
    });
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
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