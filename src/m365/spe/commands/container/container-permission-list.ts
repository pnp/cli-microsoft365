import { cli } from '../../../../cli/cli.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { spe } from '../../../../utils/spe.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  containerId: z.string().alias('i').optional(),
  containerName: z.string().alias('n').optional(),
  containerTypeId: z.uuid().optional(),
  containerTypeName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerPermissionListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists permissions of a SharePoint Embedded Container';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'userPrincipalName', 'roles'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((opts: Options) => [opts.containerId, opts.containerName].filter(value => value !== undefined).length === 1, {
        message: 'Specify either id or name, but not both.'
      })
      .refine((options: Options) => !options.containerName || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
      })
      .refine((options: Options) => options.containerName || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
        error: 'Options containerTypeId and containerTypeName are only required when retrieving a container by name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const containerId = await this.resolveContainerId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving permissions of a SharePoint Embedded Container with id '${containerId}'...`);
      }

      const containerPermission = await odata.getAllItems<any>(`${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(containerPermission);
      }
      else {
        await logger.log(containerPermission.map(i => {
          return {
            id: i.id,
            roles: i.roles.join(','),
            userPrincipalName: i.grantedToV2.user.userPrincipalName
          };
        }));
      }
    }
    catch (err: any) {
      if (err instanceof CommandError) {
        throw err;
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async resolveContainerId(options: Options, logger: Logger): Promise<string> {
    if (options.containerId) {
      return options.containerId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Resolving container id from name '${options.containerName}'...`);
    }

    const containerTypeId = await this.getContainerTypeId(options, logger);
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

export default new SpeContainerPermissionListCommand();
