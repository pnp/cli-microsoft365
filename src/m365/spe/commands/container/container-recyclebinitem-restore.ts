import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { SpeContainer, spe } from '../../../../utils/spe.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    containerTypeId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    containerTypeName: z.string().optional(),
    id: z.string().optional(),
    name: z.string().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerRecycleBinItemRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted container';
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
        message: 'Options containerTypeId and containerTypeName are only required when restoring a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const containerId = await this.getContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Restoring deleted container with ID '${containerId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/deletedContainers/${containerId}/restore`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

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
      await logger.logToStderr(`Retrieving container with name '${options.name}'...`);
    }

    const containerTypeId = await this.getContainerTypeId(options, logger);

    const containers = await odata.getAllItems<SpeContainer>(`${this.resource}/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`);
    const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === options.name!.toLowerCase());

    if (matchingContainers.length === 0) {
      throw new Error(`The specified container '${options.name}' does not exist.`);
    }

    if (matchingContainers.length > 1) {
      const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
      const container = await cli.handleMultipleResultsFound<SpeContainer>(`Multiple containers with name '${options.name}' found.`, containerKeyValuePair);
      return container.id;
    }

    return matchingContainers[0].id;
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

export default new SpeContainerRecycleBinItemRestoreCommand();