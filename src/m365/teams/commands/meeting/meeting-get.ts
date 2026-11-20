import { Event } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The userId value must be a valid GUID.'
    }).optional().alias('u'),
  userName: z.string().optional().alias('n'),
  email: z.string().optional(),
  joinUrl: z.string().alias('j')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_GET;
  }

  public get description(): string {
    return 'Gets specified meeting details';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName && !args.options.email) {
        this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`);
      }
    }
    else {
      if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
        this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`);
      }
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving meeting for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
    }

    try {
      let requestUrl = `${this.resource}/v1.0/`;

      if (isAppOnlyAccessToken) {
        requestUrl += 'users/';

        const userId = await this.getUserId(args.options);
        requestUrl += userId;
      }
      else {
        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings?$filter=JoinWebUrl eq '${formatting.encodeQueryParameter(args.options.joinUrl)}'`;

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Event[] }>(requestOptions);

      if (res.value.length > 0) {
        await logger.log(res.value[0]);
      }
      else {
        throw `The specified meeting was not found`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(options: Options): Promise<string> {
    if (options.userId) {
      return options.userId;
    }

    if (options.userName) {
      return entraUser.getUserIdByUpn(options.userName);
    }

    return entraUser.getUserIdByEmail(options.email!);
  }
}

export default new TeamsMeetingGetCommand();