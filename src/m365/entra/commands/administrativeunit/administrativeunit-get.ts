import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import request, { CliRequestOptions } from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { entraAdministrativeUnit } from "../../../../utils/entraAdministrativeUnit.js";
import { globalOptionsZod } from "../../../../Command.js";
import { zod } from "../../../../utils/zod.js";

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().uuid().optional()),
    displayName: zod.alias('n', z.string().optional()),
    properties: zod.alias('p', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_GET;
  }

  public get description(): string {
    return 'Gets information about a specific administrative unit';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.id, options.displayName].filter(Boolean).length === 1, {
        message: 'Specify either id or displayName'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnit: AdministrativeUnit;

    try {
      if (args.options.id) {
        administrativeUnit = await this.getAdministrativeUnitById(args.options.id, args.options.properties);
      }
      else {
        administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName!);
      }

      await logger.log(administrativeUnit);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  async getAdministrativeUnitById(id: string, properties?: string): Promise<AdministrativeUnit> {
    const queryParameters: string[] = [];

    if (properties) {
      const allProperties = properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/administrativeUnits/${id}${queryString}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<AdministrativeUnit>(requestOptions);
  }
}

export default new EntraAdministrativeUnitGetCommand();