import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { EngageRole } from './EngageRole.js';

class VivaEngageRoleListCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_ROLE_LIST;
  }

  public get description(): string {
    return 'Lists all Viva Engage roles';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting all Viva Engage roles...');
    }

    try {
      const results = await odata.getAllItems<EngageRole>(`${this.resource}/beta/employeeExperience/roles`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageRoleListCommand();