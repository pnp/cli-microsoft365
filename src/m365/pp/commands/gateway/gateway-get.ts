import { Logger } from "../../../../cli/Logger.js";
import { CommandArgs } from "../../../../Command.js";
import request from "../../../../request.js";
import { formatting } from "../../../../utils/formatting.js";
import { validation } from "../../../../utils/validation.js";
import PowerBICommand from "../../../base/PowerBICommand.js";
import commands from "../../commands.js";

class PpGatewayGetCommand extends PowerBICommand {
  public get name(): string {
    return commands.GATEWAY_GET;
  }

  public get description(): string {
    return "Get information about the specified gateway.";
  }

  constructor() {
    super();
    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift({
      option: "-i, --id <id>"
    });
  }

  #initValidators(): void {
    this.validators.push(async (args: CommandArgs) => {
      if (!validation.isValidGuid(args.options.id as string)) {
        return `${args.options.id} is not a valid GUID`;
      }
      return true;
    });
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