import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import request, { CliRequestOptions } from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { aadAdministrativeUnit } from "../../../../utils/aadAdministrativeUnit.js";

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
}

class AadAdministrativeUnitGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_GET;
  }

  public get description(): string {
    return 'Gets information about a specific administrative unit';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName'] });
  }

  #initTypes(): void {
    this.types.string.push('displayName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnit: AdministrativeUnit;

    try {
      if (args.options.id) {
        administrativeUnit = await this.getAdministrativeUnitById(args.options.id);
      }
      else {
        administrativeUnit = await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.displayName!);
      }

      await logger.log(administrativeUnit);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  async getAdministrativeUnitById(id: string): Promise<AdministrativeUnit> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/administrativeUnits/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<AdministrativeUnit>(requestOptions);
  }
}

export default new AadAdministrativeUnitGetCommand();