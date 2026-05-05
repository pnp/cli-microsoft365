import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  name: z.string().optional(),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarGroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDARGROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes a calendar group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.id || options.name, {
        error: 'Specify either id or name.'
      })
      .refine(options => !(options.id && options.name), {
        error: 'Specify either id or name, but not both.'
      })
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeCalendarGroup = async (): Promise<void> => {
      try {
        const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }

        let endpoint: string;
        let graphUserId: string;

        if (args.options.userId || args.options.userName) {
          graphUserId = (args.options.userId ?? args.options.userName)!;
        }
        else {
          graphUserId = accessToken.getUserIdFromAccessToken(token);
        }

        endpoint = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(graphUserId)}')`;

        let calendarGroupId = args.options.id;

        if (args.options.name) {
          if (this.verbose) {
            await logger.logToStderr(`Retrieving calendar group by name '${args.options.name}'...`);
          }

          const calendarGroupResult = await calendarGroup.getUserCalendarGroupByName(graphUserId, args.options.name);
          calendarGroupId = calendarGroupResult.id!;
        }

        if (this.verbose) {
          await logger.logToStderr(`Removing calendar group '${calendarGroupId}'...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${endpoint}/calendarGroups/${calendarGroupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeCalendarGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove calendar group '${args.options.id || args.options.name}'?` });

      if (result) {
        await removeCalendarGroup();
      }
    }
  }
}

export default new OutlookCalendarGroupRemoveCommand();
