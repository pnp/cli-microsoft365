import { DirectoryObject } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import { aadAdministrativeUnit } from "../../../../utils/aadAdministrativeUnit.js";
import request, { CliRequestOptions } from "../../../../request.js";

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  administrativeUnitId?: string;
  administrativeUnitName?: string;
  properties?: string;
}

interface DirectoryObjectEx extends DirectoryObject {
  '@odata.context'?: string;
  '@odata.type'?: string;
  type: string;
}

class AadAdministrativeUnitMemberGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_GET;
  }

  public get description(): string {
    return 'Retrieve a specific member (user, group, or device) of an administrative unit';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
        administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: 'd, --name [name]'
      },
      {
        option: '-u, --administrativeUnitId [administrativeUnitId]'
      },
      {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
      },
      {
        option: '-p, --properties [properties]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId as string)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnitId = args.options.administrativeUnitId;
    let memberId = args.options.id;

    try {
      if (args.options.administrativeUnitName) {
        administrativeUnitId = (await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id!;
      }

      if (args.options.name) {
        memberId = await aadAdministrativeUnit.getMemberIdByName(args.options.name);
      }

      const url = this.getRequestUrl(administrativeUnitId!, memberId!, args.options);

      const requestOptions: CliRequestOptions = {
        url: url,
        headers: {
          accept: 'application/json;odata.metadata=minimal'
        },
        responseType: 'json'
      };

      const result = await request.get<DirectoryObjectEx>(requestOptions);
      const odataType = result['@odata.type'];

      if (odataType) {
        result.type = odataType.replace('#microsoft.graph.', '');
      }

      delete result['@odata.type'];
      delete result['@odata.context'];
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(administrativeUnitId: string, memberId: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));
      const expandProperties = allProperties.filter(prop => prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      if (expandProperties.length > 0) {
        const fieldExpands = expandProperties.map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`);
        queryParameters.push(`$expand=${fieldExpands.join(',')}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    return `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${memberId}${queryString}`;
  }
}

export default new AadAdministrativeUnitMemberGetCommand();