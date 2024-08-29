import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Community } from './Community.js';

class VivaEngageCommunityListCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_LIST;
  }

  public get description(): string {
    return 'Lists Viva Engage communities';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'privacy'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting the list of Viva Engage communities');
    }

    try {
      const results = await odata.getAllItems<Community>(`${this.resource}/v1.0/employeeExperience/communities`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageCommunityListCommand();