import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { spe } from '../../../../utils/spe.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  containerId: z.string().optional(),
  containerName: z.string().alias('n').optional(),
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional(),
  roles: z.string().alias('r').transform((value) => value.split(',')).pipe(z.enum(['reader', 'writer', 'manager', 'owner']).array())
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerPermissionSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_PERMISSION_SET;
  }

  public get description(): string {
    return 'Updates SharePoint Embedded Container permission';
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
        error: 'Options containerTypeId and containerTypeName are only required when adding permissions to a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const containerId = await this.getContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding permissions to container with ID '${containerId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          roles: args.options.roles
        }
      };

      await request.patch(requestOptions);
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

export default new SpeContainerPermissionSetCommand();