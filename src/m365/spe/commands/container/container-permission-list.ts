import { cli } from '../../../../cli/cli.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { spe } from '../../../../utils/spe.js';

const options = globalOptionsZod
  .extend({
    containerId: zod.alias('i', z.string().optional()),
    containerName: zod.alias('n', z.string().optional())
  })
  .strict();
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

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema.refine((opts: Options) => [opts.containerId, opts.containerName].filter(value => value !== undefined).length === 1, {
      message: 'Specify either containerId or containerName, but not both.'
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

    try {
      return await spe.getContainerIdByName(options.containerName!);
    }
    catch (error: any) {
      this.handleRejectedODataJsonPromise(error);
      throw error;
    }
  }
}

export default new SpeContainerPermissionListCommand();
