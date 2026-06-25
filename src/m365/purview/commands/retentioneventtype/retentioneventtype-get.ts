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
  }).alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewRetentionEventTypeGetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_GET;
  }

  public get description(): string {
    return 'Retrieve the specified retention event type';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving retention event type with id ${args.options.id}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: any = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionEventTypeGetCommand();