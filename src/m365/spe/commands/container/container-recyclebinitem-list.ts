import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { spe } from '../../../../utils/spe.js';
import { odata } from '../../../../utils/odata.js';

const options = globalOptionsZod
  .extend({
    containerTypeId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    containerTypeName: z.string().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists deleted containers of a specific container type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine((options: Options) => [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: containerTypeId or containerTypeName.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const containerTypeId = await this.getContainerTypeId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving deleted containers of container type with ID '${containerTypeId}'...`);
      }

      const deletedContainers = await odata.getAllItems<any>(`${this.resource}/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`);
      await logger.log(deletedContainers);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContainerTypeId(options: Options, logger: Logger): Promise<string> {
    if (options.containerTypeId) {
      return options.containerTypeId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving container type id for container type '${options.containerTypeName}'...`);
    }

    return spe.getContainerTypeIdByName(options.containerTypeName!);
  }
}

export default new SpeContainerRecycleBinItemListCommand();