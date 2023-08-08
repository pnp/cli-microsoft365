import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class PurviewRetentionEventTypeListCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_LIST;
  }

  public get description(): string {
    return 'Get a list of retention event types';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'createdDateTime'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const items = await odata.getAllItems(`${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionEventTypeListCommand();