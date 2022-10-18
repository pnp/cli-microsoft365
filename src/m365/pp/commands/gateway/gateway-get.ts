import { Logger } from "../../../../cli/Logger";
import { CommandArgs } from "../../../../Command";
import request from "../../../../request";
import { validation } from "../../../../utils/validation";
import PowerBICommand from "../../../base/PowerBICommand";
import commands from "../../commands";

class PpGatewayGetCommand extends PowerBICommand {
  public get name(): string {
    return commands.GATEWAY_GET;
  }

  public get description(): string {
    return "Returns the specified gateway.";
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
      logger.log(gateway);
    } catch (error) {
      this.handleRejectedODataJsonPromise(error);
    }
  }

  private getGateway(gatewayId: string): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorg/gateways/${encodeURIComponent(
        gatewayId
      )}`,
      headers: {
        accept: "application/json;odata.metadata=none",
      },
      responseType: "json"
    };

    return request.get<any>(requestOptions);
  }
}

module.exports = new PpGatewayGetCommand();
