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
  type?: string;
  properties?: string;
  filter?: string;
}

interface DirectoryObjectEx extends DirectoryObject {
  '@odata.type'?: string;
  type: string;
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
        type: typeof args.options.type !== 'undefined',
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
        option: '-t, --type [type]',
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

        if (args.options.type) {
          if (['user', 'group', 'device'].every(type => type !== args.options.type)) {
            return `${args.options.type} is not a valid type value. Allowed values user|group|device`;
          }
        }

        if (args.options.filter && !args.options.type) {
          return 'Filter can be specified only if type is set';
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

      let results;
      const endpoint = this.getRequestUrl(administrativeUnitId!, args.options);

      if (args.options.type) {
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
        results = await odata.getAllItems<DirectoryObjectEx>(endpoint, 'minimal');

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

  private getRequestUrl(administrativeUnitId: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
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

    if (options.filter) {
      queryParameters.push(`$filter=${options.filter}`);
      queryParameters.push('$count=true');
    }

    let query = '';

    for (let i = 0; i < queryParameters.length; i++) {
      query += i === 0 ? '?' : '&';
      query += queryParameters[i];
    }

    let endpoint;

    if (options.type) {
      endpoint = `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.${options.type}${query}`;
    }
    else {
      endpoint = `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members${query}`;
    }

    return endpoint;
  }
}

export default new AadAdministrativeUnitMemberListCommand();