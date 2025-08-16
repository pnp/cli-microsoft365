import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import GraphApplicationCommand from '../../../base/GraphApplicationCommand.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    id: z.string()
      .refine((val) => validation.isValidGuid(val), {
        message: 'Invalid GUID.'
      }).optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsCallRecordGetCommand extends GraphApplicationCommand {
  public get name(): string {
    return commands.CALLRECORD_GET;
  }

  public get description(): string {
    return 'Gets a specific Teams call';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const callRecordId = args.options.id;
      if (this.verbose) {
        await logger.logToStderr(`Retrieving call record {callRecordId}...`);
      }

      // only one relationship can be expanded at a time
      let requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/communications/callRecords/${callRecordId}?$expand=sessions($expand=segments)`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const callRecordPart1 = await request.get<any>(requestOptions);

      requestOptions = {
        url: `${this.resource}/v1.0/communications/callRecords/${callRecordId}?$select=id&$expand=participants_v2`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const callRecordPart2 = await request.get<any>(requestOptions);

      const callRecord = { ...callRecordPart1, ...callRecordPart2 };

      await logger.log(callRecord);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsCallRecordGetCommand();