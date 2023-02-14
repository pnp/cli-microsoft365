import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
      logger.logToStderr(`Retrieving the commercial subscriptions that an organization has acquired`);
    }

    try {
      const items = await odata.getAllItems<any>(`${this.resource}/v1.0/subscribedSkus`);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadLicenseListCommand();