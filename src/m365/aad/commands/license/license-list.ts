import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class AadLicenseListCommand extends GraphCommand {
  public get name(): string {
    return commands.LICENSE_LIST;
  }

  public get description(): string {
    return 'Lists commercial subscriptions that an organization has acquired';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'skuId', 'skuPartNumber'];
  }


  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the commercial subscriptions that an organization has acquired`);
    }

    try {
      const items = await odata.getAllItems<any>(`${this.resource}/v1.0/subscribedSkus`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadLicenseListCommand();