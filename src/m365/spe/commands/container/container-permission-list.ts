import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';

const options = globalOptionsZod
  .extend({
    containerId: zod.alias('i', z.string())
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving permissions of a SharePoint Embedded Container with id '${args.options.containerId}'...`);
      }

      const containerPermission = await odata.getAllItems<any>(`${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(args.options.containerId)}/permissions`);
      await logger.log(containerPermission.map(i => {
        return {
          id: i.id,
          roles: i.roles.join(','),
          userPrincipalName: i.grantedToV2.user.userPrincipalName
        };
      }));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerPermissionListCommand();