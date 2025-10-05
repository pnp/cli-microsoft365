import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { spe } from '../../../../utils/spe.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n'),
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional(),
  recycle: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_REMOVE;
  }

  public get description(): string {
    return 'Removes a container';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((options: Options) => [options.id, options.name].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id or name.'
      })
      .refine((options: Options) => !options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
      })
      .refine((options: Options) => options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
        error: 'Options containerTypeId and containerTypeName are only required when deleting a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove container '${args.options.id || args.options.name}'${!args.options.recycle ? ' permanently' : ''}?` });

      if (!result) {
        return;
      }
    }

    try {
      const containerId = await this.getContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Removing container with ID '${containerId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${containerId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      if (args.options.recycle) {
        await request.delete(requestOptions);
        return;
      }

      // Container should be permanently deleted
      requestOptions.url += '/permanentDelete';
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContainerId(options: Options, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const containerTypeId = await this.getContainerTypeId(options, logger);

    if (this.verbose) {
      await logger.logToStderr(`Getting container ID for container with name '${options.name}'...`);
    }

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

export default new SpeContainerRemoveCommand();