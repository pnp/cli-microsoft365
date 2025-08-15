import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from "../../../../request.js";
import { zod } from '../../../../utils/zod.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";

const options = globalOptionsZod
  .extend({
    displayName: zod.alias('n', z.string()),
    description: zod.alias('d', z.string().optional()),
    hiddenMembership: z.boolean().optional()
  })
  .passthrough();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_ADD;
  }

  public get description(): string {
    return 'Creates an administrative unit';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestBody = {
      description: args.options.description,
      displayName: args.options.displayName,
      visibility: args.options.hiddenMembership ? 'HiddenMembership' : null
    };

    this.addUnknownOptionsToPayload(requestBody, args.options);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/administrativeUnits`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: requestBody
    };

    try {
      const administrativeUnit = await request.post<AdministrativeUnit>(requestOptions);

      await logger.log(administrativeUnit);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAdministrativeUnitAddCommand();