import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { SpeContainer } from '../../../../utils/spe.js';
import { zod } from '../../../../utils/zod.js';

const options = globalOptionsZod.extend({
  id: zod.alias('i', z.string().optional()),
  name: zod.alias('n', z.string().optional())
}).strict();

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

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema.refine((opts: Options) => [opts.id, opts.name].filter(value => value !== undefined).length === 1, {
      message: 'Specify either id or name, but not both.'
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
      if (err instanceof CommandError) {
        throw err;
      }

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

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/storage/fileStorage/containers`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<{ value?: SpeContainer[] }>(requestOptions);
      const container = response.value?.find(item => item.displayName === options.name);

      if (!container) {
        throw new CommandError(`Container with name '${options.name}' not found.`);
      }

      return container.id;
    }
    catch (error: any) {
      if (error instanceof CommandError) {
        throw error;
      }

      throw error;
    }
  }
}

export default new SpeContainerGetCommand();
