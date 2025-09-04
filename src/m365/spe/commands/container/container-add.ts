import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { spe } from '../../../../utils/spe.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string()),
    description: zod.alias('d', z.string()).optional(),
    containerTypeId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    containerTypeName: z.string().optional(),
    ocrEnabled: z.boolean().optional(),
    itemMajorVersionLimit: z.number()
      .refine(numb => validation.isValidPositiveInteger(numb), numb => ({
        message: `'${numb}' is not a valid positive integer.`
      })).optional(),
    itemVersioningEnabled: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_ADD;
  }

  public get description(): string {
    return 'Creates a new container';
  }

  public get schema(): z.ZodTypeAny {
    return options;
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
        await logger.logToStderr(`Creating container with name '${args.options.name}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          displayName: args.options.name,
          description: args.options.description,
          containerTypeId: containerTypeId,
          settings: {
            isOcrEnabled: args.options.ocrEnabled,
            itemMajorVersionLimit: args.options.itemMajorVersionLimit,
            isItemVersioningEnabled: args.options.itemVersioningEnabled
          }
        }
      };

      const container = await request.post<any>(requestOptions);
      await logger.log(container);
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
      await logger.logToStderr(`Getting container type with name '${options.containerTypeName}'...`);
    }

    return spe.getContainerTypeIdByName(options.containerTypeName!);
  }
}

export default new SpeContainerAddCommand();