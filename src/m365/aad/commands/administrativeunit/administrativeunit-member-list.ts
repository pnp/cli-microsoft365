import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { aadAdministrativeUnit } from '../../../../utils/aadAdministrativeUnit.js';
import { validation } from '../../../../utils/validation.js';
import { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  administrativeUnitId?: string;
  administrativeUnitName?: string;
  memberType?: string;
  properties?: string;
  filter?: string;
}

interface DirectoryObjectEx extends DirectoryObject {
  '@odata.type'?: string;
  type: string;
}

interface GraphQueryParameters {
  /**
   * List of properties separated by a comma. Properties without a slash are used in $select query parameter. 
   * Propeties with a slash are used in $expand query parameter.
   */
  properties?: string;

  /** Filter expression used in $filter query parameter.*/
  filter?: string;

  /** If specified then $count=true is included.*/
  count?: boolean;
}

class AadAdministrativeUnitMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_LIST;
  }

  public get description(): string {
    return 'Retrieves members (users, groups, or devices) of an administrative unit.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
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
        memberType: typeof args.options.memberType !== 'undefined',
        properties: typeof args.options.properties !== 'undefined',
        filter: typeof args.options.filter !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --administrativeUnitId [administrativeUnitId]'
      },
      {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
      },
      {
        option: '-m, --memberType [memberType]',
        autocomplete: ['user', 'group', 'device']
      },
      {
        option: '-p, --properties [properties]'
      },
      {
        option: '-f, --filter [filter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId as string)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.memberType) {
          if (['user', 'group', 'device'].indexOf(args.options.memberType) === -1) {
            return `${args.options.memberType} is not a valid memberType value. Allowed values user|group|device`;
          }
        }

        if (args.options.filter && !args.options.memberType) {
          return 'Filter can be specified only if memberType is set';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnitId = args.options.administrativeUnitId;

    try {
      if (args.options.administrativeUnitName) {
        administrativeUnitId = (await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
      }

      const queryInputParameters = { properties: args.options.properties, filter: args.options.filter, count: false };
      let results;

      if (args.options.memberType) {
        if (args.options.filter) {
          queryInputParameters.count = true;
        }

        const query = this.createGraphQuery(queryInputParameters);
        const endpoint = `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.${args.options.memberType}${query}`;

        if (args.options.filter) {
          // While using the filter, we need to specify the ConsistencyLevel header.
          // Can be refactored when the header is no longer necessary.
          const requestOptions: CliRequestOptions = {
            url: endpoint,
            headers: {
              accept: 'application/json;odata.metadata=none',
              ConsistencyLevel: 'eventual'
            },
            responseType: 'json'
          };
          results = await odata.getAllItems<DirectoryObject>(requestOptions);
        }
        else {
          results = await odata.getAllItems<DirectoryObject>(endpoint);
        }
      }
      else {
        const query = this.createGraphQuery(queryInputParameters);
        results = await odata.getAllItems<DirectoryObjectEx>(`${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members${query}`, 'minimal');

        results.forEach(c => {
          const odataType = c['@odata.type'];

          if (odataType) {
            c.type = odataType.replace('#microsoft.graph.', '');
          }

          delete c['@odata.type'];
        });
      }

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Create a query for a request to the Graph API
   * @param parameters Parameters to be applied to the query.
   * @returns Query with applied parameters.
  */
  public createGraphQuery(parameters: GraphQueryParameters): string {
    const queryParameters: string[] = [];

    if (parameters.properties) {
      const allProperties = parameters.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      const expandProperties = allProperties.filter(prop => prop.includes('/'));

      let fieldExpand: string = '';

      expandProperties.forEach(p => {
        if (fieldExpand.length > 0) {
          fieldExpand += ',';
        }

        fieldExpand += `${p.split('/')[0]}($select=${p.split('/')[1]})`;
      });

      if (fieldExpand.length > 0) {
        queryParameters.push(`$expand=${fieldExpand}`);
      }
    }

    if (parameters.filter) {
      queryParameters.push(`$filter=${parameters.filter}`);
    }

    if (parameters.count) {
      queryParameters.push('$count=true');
    }

    let query = '';

    for (let i = 0; i < queryParameters.length; i++) {
      query += i === 0 ? '?' : '&';
      query += queryParameters[i];
    }

    return query;
  }
}

export default new AadAdministrativeUnitMemberListCommand();