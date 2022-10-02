import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class BookingBusinessListCommand extends GraphCommand {
  public get name(): string {
    return commands.BUSINESS_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Bookings businesses that are created for the tenant.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/solutions/bookingBusinesses`;

    try {
      const items = await odata.getAllItems<BookingBusiness>(endpoint);
      logger.log(items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new BookingBusinessListCommand();