import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spe, SpeContainer } from '../../../../utils/spe.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i').optional(),
  name: z.string().alias('n').optional(),
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional()
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_GET;
  }

  public get description(): string {
    return 'Gets a container of a specific container type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((opts: Options) => [opts.id, opts.name].filter(value => value !== undefined).length === 1, {
        message: 'Specify either id or name, but not both.'
      })
      .refine((options: Options) => !options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
      })
      .refine((options: Options) => options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
        error: 'Options containerTypeId and containerTypeName are only required when retrieving a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const containerId = await this.resolveContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Getting a container with id '${containerId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${containerId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<SpeContainer>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async resolveContainerId(options: Options, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Resolving container id from name '${options.name}'...`);
    }

    const containerTypeId = await this.getContainerTypeId(options, logger);
    return spe.getContainerIdByName(containerTypeId, options.name!);
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

export default new SpeContainerGetCommand();
