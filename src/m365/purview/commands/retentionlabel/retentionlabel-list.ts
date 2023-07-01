import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class PurviewRetentionLabelListCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_LIST;
  }

  public get description(): string {
    return 'Get a list of retention labels';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isInUse'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const items = await odata.getAllItems(`${this.resource}/beta/security/labels/retentionLabels`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionLabelListCommand();