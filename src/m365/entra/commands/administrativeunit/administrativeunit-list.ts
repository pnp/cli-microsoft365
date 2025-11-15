import { z } from 'zod';
import { AdministrativeUnit } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  properties: z.string().optional().alias('p')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitListCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of administrative units';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'visibility'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const queryParameters: string[] = [];

    if (args.options.properties) {
      const allProperties = args.options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    try {
      const results = await odata.getAllItems<AdministrativeUnit>(`${this.resource}/v1.0/directory/administrativeUnits${queryString}`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAdministrativeUnitListCommand();