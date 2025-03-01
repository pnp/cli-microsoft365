import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { UnifiedRbacResourceAction } from '@microsoft/microsoft-graph-types';

const options = globalOptionsZod
  .extend({
    resourceNamespace: zod.alias('n', z.string()),
    privileged: zod.alias('p', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface UnifiedRbacResourceActionExt extends UnifiedRbacResourceAction {
  isPrivileged?: boolean;
}

class EntraRolePermissionListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEPERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Entra ID role permissions';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'actionVerb', 'isPrivileged'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting Microsoft Entra ID role permissions...');
    }

    try {
      const queryString = args.options.privileged ? '?$filter=isPrivileged eq true' : '';
      const url = `${this.resource}/beta/roleManagement/directory/resourceNamespaces/${args.options.resourceNamespace}/resourceActions${queryString}`;
      const results = await odata.getAllItems<UnifiedRbacResourceActionExt>(url);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraRolePermissionListCommand();
