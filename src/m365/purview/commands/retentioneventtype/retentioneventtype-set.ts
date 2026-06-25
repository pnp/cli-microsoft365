import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).alias('i'),
  description: z.string().optional().alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewRetentionEventTypeSetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_SET;
  }

  public get description(): string {
    return 'Update a retention event type';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.description, {
        message: 'Specify at least one option to update.',
        params: {
          customCode: 'required'
        }
      }) as any;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.log(`Updating retention event type with id ${args.options.id}`);
    }

    try {
      const requestBody = {
        description: args.options.description
      };

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${args.options.id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json',
        data: requestBody
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionEventTypeSetCommand();