import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const businessId = await this.getBusinessId(args.options);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(businessId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const business = await request.get<BookingBusiness>(requestOptions);
      logger.log(business);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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