import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { SpeContainer, spe } from '../../../../utils/spe.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_LIST;
  }

  public get description(): string {
    return 'Lists containers of a specific Container Type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.containerTypeId, opts.containerTypeName].filter(o => o !== undefined).length === 1, {
        message: 'Specify one of the following options: containerTypeId, containerTypeName.',
        params: {
          customCode: 'optionSet',
          options: ['containerTypeId', 'containerTypeName']
        }
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'containerTypeId', 'createdDateTime'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of Containers...`);
      }

      const containerTypeId = await this.getContainerTypeId(logger, args.options);
      const allContainers = await odata.getAllItems<SpeContainer>(`${this.resource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${formatting.encodeQueryParameter(containerTypeId)}`);
      await logger.log(allContainers);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getContainerTypeId(logger: Logger, options: Options): Promise<string> {
    if (options.containerTypeId) {
      return options.containerTypeId;
    }

    return spe.getContainerTypeIdByName(options.containerTypeName!);
  }
}

export default new SpeContainerListCommand();