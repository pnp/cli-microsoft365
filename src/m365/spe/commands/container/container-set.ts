import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  newName: z.string().optional(),
  description: z.string().optional(),
  isOcrEnabled: z.boolean().optional(),
  isItemVersioningEnabled: z.boolean().optional(),
  itemMajorVersionLimit: z.number()
    .refine(numb => validation.isValidPositiveInteger(numb), {
      error: e => `'${e.input}' is not a valid positive integer.`
    }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_SET;
  }

  public get description(): string {
    return 'Updates a SharePoint Embedded container';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(o => o.newName !== undefined || o.description !== undefined || o.isOcrEnabled !== undefined || o.isItemVersioningEnabled !== undefined || o.itemMajorVersionLimit !== undefined, {
        error: 'Specify at least one of newName, description, isOcrEnabled, isItemVersioningEnabled, or itemMajorVersionLimit.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating container '${args.options.id}'...`);
    }

    try {
      const data: any = {};

      if (args.options.newName !== undefined) {
        data.displayName = args.options.newName;
      }

      if (args.options.description !== undefined) {
        data.description = args.options.description;
      }

      const settings: any = {};
      if (args.options.isOcrEnabled !== undefined) {
        settings.isOcrEnabled = args.options.isOcrEnabled;
      }

      if (args.options.isItemVersioningEnabled !== undefined) {
        settings.isItemVersioningEnabled = args.options.isItemVersioningEnabled;
      }

      if (args.options.itemMajorVersionLimit !== undefined) {
        settings.itemMajorVersionLimit = args.options.itemMajorVersionLimit;
      }

      if (Object.keys(settings).length > 0) {
        data.settings = settings;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data
      };

      const container = await request.patch<any>(requestOptions);
      await logger.log(container);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerSetCommand();
