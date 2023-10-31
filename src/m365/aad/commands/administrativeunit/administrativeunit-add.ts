import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import { Logger } from "../../../../cli/Logger.js";
import request, { CliRequestOptions } from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description?: string;
  hiddenMembership?: boolean;
}

class AadAdministrativeUnitAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_ADD;
  }

  public get description(): string {
    return 'Creates an administrative unit';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        hiddenMembership: !!args.options.hiddenMembership
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--hiddenMembership'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/administrativeUnits`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.displayName,
        visibility: args.options.hiddenMembership ? 'HiddenMembership' : null
      }
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

export default new AadAdministrativeUnitAddCommand();