import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0/solutions/bookingBusinesses`;

    odata
      .getAllItems<BookingBusiness>(endpoint)
      .then((items): Promise<BookingBusiness[]> => {
        return Promise.resolve(items);
      })
      .then((items: BookingBusiness[]): void => {
        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new BookingBusinessListCommand();