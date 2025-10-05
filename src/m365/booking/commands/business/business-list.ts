import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({ ...globalOptionsZod.shape });

class BookingBusinessListCommand extends GraphCommand {
  public get name(): string {
    return commands.BUSINESS_LIST;
  }

  public get description(): string {
    return 'Lists all Microsoft Bookings businesses that are created for the tenant.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/solutions/bookingBusinesses`;

    try {
      const items = await odata.getAllItems<BookingBusiness>(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new BookingBusinessListCommand();