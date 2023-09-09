import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Cli } from '../../../../cli/Cli.js';

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
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const businessId = await this.getBusinessId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(businessId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const business = await request.get<BookingBusiness>(requestOptions);
      await logger.log(business);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getBusinessId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/solutions/bookingBusinesses`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: BookingBusiness[] }>(requestOptions);

    const name = options.name as string;
    const bookingBusinesses: BookingBusiness[] | undefined = response.value.filter(val => val.displayName?.toLocaleLowerCase() === name.toLocaleLowerCase());

    if (!bookingBusinesses.length) {
      throw `The specified business with name ${options.name} does not exist.`;
    }

    if (bookingBusinesses.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', bookingBusinesses);
      const result = await Cli.handleMultipleResultsFound<BookingBusiness>(`Multiple businesses with name '${options.name}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return bookingBusinesses[0].id!;
  }
}

export default new BookingBusinessGetCommand();