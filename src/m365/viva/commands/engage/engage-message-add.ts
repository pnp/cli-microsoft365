import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  body: z.string(),
  repliedToId: z.coerce.number().optional(),
  directToUserIds: z.string().optional(),
  groupId: z.coerce.number().optional(),
  networkId: z.coerce.number().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageMessageAddCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_MESSAGE_ADD;
  }

  public get description(): string {
    return 'Posts a Viva Engage network message on behalf of the current user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.groupId !== undefined || opts.directToUserIds !== undefined || opts.repliedToId !== undefined, {
        message: 'You must either specify groupId, repliedToId or directToUserIds',
        params: { customCode: 'required' }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${this.resource}/v1/messages.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        body: args.options.body,
        replied_to_id: args.options.repliedToId,
        direct_to_user_ids: args.options.directToUserIds,
        group_id: args.options.groupId,
        network_id: args.options.networkId
      }
    };

    try {
      const res: any = await request.post(requestOptions);
      let result = null;
      if (res.messages && res.messages.length === 1) {
        result = res.messages[0];
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageMessageAddCommand();
