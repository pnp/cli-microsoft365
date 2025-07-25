import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().uuid().optional()),
    displayName: zod.alias('n', z.string().optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_REMOVE;
  }
  public get description(): string {
    return 'Removes an administrative unit';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => options.id || options.displayName, {
        message: 'Specify either id or displayName'
      })
      .refine(options => !(options.id && options.displayName), {
        message: 'Specify either id or displayName but not both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeAdministrativeUnit = async (): Promise<void> => {
      try {
        let administrativeUnitId = args.options.id;

        if (args.options.displayName) {
          const administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName);
          administrativeUnitId = administrativeUnit.id;
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeAdministrativeUnit();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove administrative unit '${args.options.id || args.options.displayName}'?` });

      if (result) {
        await removeAdministrativeUnit();
      }
    }
  }
}

export default new EntraAdministrativeUnitRemoveCommand();