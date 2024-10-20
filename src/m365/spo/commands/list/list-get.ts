import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DefaultTrimModeType, ListInstance, VersionPolicy } from "./ListInstance.js";
import { ListPrincipalType } from './ListPrincipalType.js';

interface Properties {
  selectProperties: string[],
  expandProperties: string[]
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
  url?: string;
  properties?: string;
  withPermissions?: boolean;
  default?: boolean;
}

class SpoListGetCommand extends SpoCommand {
  private supportedBaseTemplates = [101, 109, 110, 111, 113, 114, 115, 116, 117, 119, 121, 122, 123, 126, 130, 175];

  public get name(): string {
    return commands.LIST_GET;
  }

  public get description(): string {
    return 'Gets information about the specific list';
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
        title: typeof args.options.title !== 'undefined',
        url: typeof args.options.url !== 'undefined',
        properties: typeof args.options.properties !== 'undefined',
        withPermissions: !!args.options.withPermissions,
        default: !!args.options.default
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--url [url]'
      },
      {
        option: '--default'
      },
      {
        option: '-p, --properties [properties]'
      },
      {
        option: '--withPermissions'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title', 'url', 'default'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information for list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    if (args.options.id) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.title) {
      requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title)}')`;
    }
    else if (args.options.url) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }
    else if (args.options.default) {
      requestUrl += `DefaultDocumentLibrary`;
    }

    const fieldsProperties: Properties = this.formatSelectProperties(args.options.properties, args.options.withPermissions);
    const queryParams: string[] = [];

    if (fieldsProperties.selectProperties.length > 0) {
      queryParams.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
    }

    if (fieldsProperties.expandProperties.length > 0) {
      queryParams.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}${queryParams.length > 0 ? `?${queryParams.join('&')}` : ''}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const listInstance = await request.get<ListInstance>(requestOptions);
      if (args.options.withPermissions) {
        listInstance.RoleAssignments.forEach(r => {
          r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
        });
      }

      if (this.supportedBaseTemplates.some(template => template === listInstance.BaseTemplate)) {
        await this.retrieveVersionPolicies(requestUrl, listInstance);
      }

      if (listInstance.VersionPolicies) {
        listInstance.VersionPolicies.DefaultTrimModeValue = DefaultTrimModeType[listInstance.VersionPolicies.DefaultTrimMode];
      }

      await logger.log(listInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private formatSelectProperties(properties: string | undefined, withPermissions: boolean | undefined): Properties {
    const selectProperties: any[] = [];
    let expandProperties: any[] = [];

    if (withPermissions) {
      expandProperties = ['HasUniqueRoleAssignments', 'RoleAssignments/Member', 'RoleAssignments/RoleDefinitionBindings', 'VersionPolicies'];
    }

    if (properties) {
      properties.split(',').forEach((property) => {
        const subparts = property.trim().split('/');
        if (subparts.length > 1) {
          expandProperties.push(subparts[0]);
        }
        selectProperties.push(property.trim());
      });
    }

    return {
      selectProperties: [...new Set(selectProperties)],
      expandProperties: [...new Set(expandProperties)]
    };
  }

  private async retrieveVersionPolicies(requestUrl: string, listInstance: ListInstance): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$select=VersionPolicies&$expand=VersionPolicies`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const { VersionPolicies } = await request.get<{ VersionPolicies: VersionPolicy }>(requestOptions);
    listInstance.VersionPolicies = VersionPolicies;
    listInstance.VersionPolicies.DefaultTrimModeValue = DefaultTrimModeType[listInstance.VersionPolicies.DefaultTrimMode];
  }
}

export default new SpoListGetCommand();