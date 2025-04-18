import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { spe } from '../../../../utils/spe.js';
import { spo } from '../../../../utils/spo.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string()).optional(),
    name: zod.alias('n', z.string()).optional(),
    containerTypeId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    containerTypeName: z.string().optional(),
    recycle: z.boolean().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

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

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine((options: Options) => [options.id, options.name].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: id or name.'
      })
      .refine((options: Options) => !options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
      })
      .refine((options: Options) => options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
        message: 'Options containerTypeId and containerTypeName are only required when deleting a container by name.'
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

    if (this.verbose) {
      await logger.logToStderr(`Getting container ID for container with name '${options.name}'...`);
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

    const adminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
    return spe.getContainerTypeIdByName(adminUrl, options.containerTypeName!);
  }
}

export default new SpeContainerRemoveCommand();