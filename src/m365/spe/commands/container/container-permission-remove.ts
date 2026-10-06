import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { spe } from '../../../../utils/spe.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  containerId: z.string().optional(),
  containerName: z.string().alias('n').optional(),
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional(),
  force: z.boolean().alias('f').optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerPermissionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_PERMISSION_REMOVE;
  }

  public get description(): string {
    return 'Removes SharePoint Embedded Container permission';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((options: Options) => [options.containerId, options.containerName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: containerId or containerName.'
      })
      .refine((options: Options) => !options.containerName || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
      })
      .refine((options: Options) => options.containerName || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
        error: 'Options containerTypeId and containerTypeName are only required when removing permissions from a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove permission '${args.options.id}' from container '${args.options.containerId || args.options.containerName}'?` });

      if (!result) {
        return;
      }
    }

    try {
      const containerId = await this.getContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Removing permissions from container with ID '${containerId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContainerId(options: Options, logger: Logger): Promise<string> {
    if (options.containerId) {
      return options.containerId;
    }

    const containerTypeId = await this.getContainerTypeId(options, logger);

    if (this.verbose) {
      await logger.logToStderr(`Getting container ID for container with name '${options.containerName}'...`);
    }

    return spe.getContainerIdByName(containerTypeId, options.containerName!);
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

export default new SpeContainerPermissionRemoveCommand();