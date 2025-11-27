import { cli } from '../../../../cli/cli.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  containerId: z.string().alias('i').optional(),
  containerName: z.string().alias('n').optional()
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

    const containers = await odata.getAllItems<{ id: string; displayName: string }>(`${this.resource}/v1.0/storage/fileStorage/containers?$select=id,displayName`);
    const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === options.containerName!.toLowerCase());

    if (matchingContainers.length === 0) {
      throw new CommandError(`The specified container '${options.containerName}' does not exist.`);
    }

    if (matchingContainers.length > 1) {
      const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
      const container = await cli.handleMultipleResultsFound<{ id: string; displayName: string }>(`Multiple containers with name '${options.containerName}' found.`, containerKeyValuePair);
      return container.id;
    }

    return matchingContainers[0].id;
  }
}

export default new SpeContainerPermissionListCommand();
