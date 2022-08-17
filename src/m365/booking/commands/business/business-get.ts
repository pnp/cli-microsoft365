import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class BookingBusinessGetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUSINESS_GET;
  }

  public get description(): string {
    return 'Retrieve the specified Microsoft Bookings business.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'businessType', 'phone', 'email', 'defaultCurrencyIso'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id [id]' },
      { option: '-n, --name [name]' }
    );
  }

  #initOptionSets(): void {
  	this.optionSets.push(['id', 'name']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {    
    this
      .getBusinessId(args.options)
      .then(businessId => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/solutions/bookingBusinesses/${encodeURIComponent(businessId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        }; 
    
        return request.get<BookingBusiness>(requestOptions);
      })
      .then(business => {
        logger.log(business);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getBusinessId(options: Options): Promise<string> {
    if (options.id) {
      return Promise.resolve(options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/solutions/bookingBusinesses`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    }; 

    return request
      .get<{ value: BookingBusiness[] }>(requestOptions)
      .then((response) => {
        const name = options.name as string;
        const bookingBusinesses: BookingBusiness[] | undefined = response.value.filter(val => val.displayName?.toLocaleLowerCase() === name.toLocaleLowerCase());

        if (!bookingBusinesses.length) {
          return Promise.reject(`The specified business with name ${options.name} does not exist.`);
        }

        if (bookingBusinesses.length > 1) {
          return Promise.reject(`Multiple businesses with name ${options.name} found. Please disambiguate: ${bookingBusinesses.map(x => x.id).join(', ')}`);
        }

        return bookingBusinesses[0].id!;
      });
  }
}

module.exports = new BookingBusinessGetCommand();