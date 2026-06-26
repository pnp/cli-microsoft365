import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import request from "../../../../request.js";
import { formatting } from "../../../../utils/formatting.js";
import { validation } from "../../../../utils/validation.js";
import PowerBICommand from "../../../base/PowerBICommand.js";
import commands from "../../commands.js";

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpGatewayGetCommand extends PowerBICommand {
  public get name(): string {
    return commands.GATEWAY_GET;
  }

  public get description(): string {
    return "Get information about the specified gateway.";
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const gateway = await this.getGateway(args.options.id);
      await logger.log(gateway);
    }
    catch (error) {
      this.handleRejectedODataJsonPromise(error);
    }
  }

  private getGateway(gatewayId: string): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorg/gateways/${formatting.encodeQueryParameter(gatewayId)}`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: "json"
    };

    return request.get<any>(requestOptions);
  }
}

export default new PpGatewayGetCommand();