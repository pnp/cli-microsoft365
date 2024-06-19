import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DefaultTrimModeType, ListInstance } from "./ListInstance.js";
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
}

class SpoListGetCommand extends SpoCommand {
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
        id: (!(!args.options.id)).toString(),
        title: (!(!args.options.title)).toString(),
        url: (!(!args.options.url)).toString(),
        properties: (!(!args.options.properties)).toString(),
        withPermissions: typeof args.options.withPermissions !== 'undefined'
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
    this.optionSets.push({ options: ['id', 'title', 'url'] });
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
      requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`;
    }
    else if (args.options.url) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const fieldsProperties: Properties = this.formatSelectProperties(args.options.properties, args.options.withPermissions);
    const queryParams: string[] = [];

    if (fieldsProperties.selectProperties.length > 0) {
      queryParams.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
    }

    if (fieldsProperties.expandProperties.length > 0) {
      queryParams.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
    }

    if (queryParams.length === 0) {
      queryParams.push(`$expand=VersionPolicies`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?${queryParams.join('&')}`,
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
}

export default new SpoListGetCommand();