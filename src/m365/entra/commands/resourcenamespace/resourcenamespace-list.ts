import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class EntraResourcenamespaceListCommand extends GraphCommand {
  public get name(): string {
    return commands.RESOURCENAMESPACE_LIST;
  }

  public get description(): string {
    return 'Get a list of the RBAC resource namespaces and their properties';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting a list of the RBAC resource namespaces and their properties...');
    }

    try {
      const results = await odata.getAllItems<{ id: string, name: string }>(`${this.resource}/beta/roleManagement/directory/resourceNamespaces`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraResourcenamespaceListCommand();