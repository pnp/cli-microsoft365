import { BookingBusiness } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().optional()),
    name: zod.alias('n', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class BookingBusinessGetCommand extends GraphCommand {
  public get name(): string {
    return commands.BUSINESS_GET;
  }

  public get description(): string {
    return 'Retrieve the specified Microsoft Bookings business.';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => options.id || options.name, {
        message: 'Specify either id or name'
      });
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
      const result = await cli.handleMultipleResultsFound<BookingBusiness>(`Multiple businesses with name '${options.name}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return bookingBusinesses[0].id!;
  }
}

export default new BookingBusinessGetCommand();