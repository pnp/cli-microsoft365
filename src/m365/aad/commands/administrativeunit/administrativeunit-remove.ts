import { AdministrativeUnit } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import request, { CliRequestOptions } from "../../../../request.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import { odata } from "../../../../utils/odata.js";
import { formatting } from "../../../../utils/formatting.js";
import { Cli } from "../../../../cli/Cli.js";


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  force?: boolean
}

class AadAdministrativeUnitRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_REMOVE;
  }
  public get description(): string {
    return 'Removes an administrative unit';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTelemetry();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: args.options.id !== 'undefined',
        displayName: args.options.displayName !== 'undefined',
        force: !!args.options.force
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
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'displayName']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID for option id.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeAdministrativeUnit = async (): Promise<void> => {

      try {
        let administrativeUnitId = args.options.id;

        if (args.options.displayName) {
          administrativeUnitId = await this.getAdministrativeUnitIdByDisplayName(args.options.displayName);
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
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove administrative unit '${args.options.id || args.options.displayName}'?`
      });

      if (result.continue) {
        await removeAdministrativeUnit();
      }
    }
  }

  async getAdministrativeUnitIdByDisplayName(displayName: string): Promise<string> {
    const administrativeUnits = await odata.getAllItems<AdministrativeUnit>(`${this.resource}/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`);

    if (administrativeUnits.length === 0) {
      throw `The specified administrative unit '${displayName}' does not exist.`;
    }

    if (administrativeUnits.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', administrativeUnits);
      const selectedAdministrativeUnit = await Cli.handleMultipleResultsFound<AdministrativeUnit>(`Multiple administrative units with name '${displayName}' found.`, resultAsKeyValuePair);
      return selectedAdministrativeUnit.id!;
    }

    return administrativeUnits[0].id!;
  }
}

export default new AadAdministrativeUnitRemoveCommand();