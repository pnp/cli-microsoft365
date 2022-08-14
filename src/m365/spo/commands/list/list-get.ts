import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from "./ListInstance";
import { ListPrincipalType } from './ListPrincipalType';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  title?: string;
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
    this.optionSets.push(['id', 'title']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`;
    }

    let propertiesSelect: string = args.options.properties ? `?$select=${encodeURIComponent(args.options.properties)}` : ``;
    propertiesSelect += args.options.withPermissions ? `${args.options.properties ? '&' : '?'}$expand=HasUniqueRoleAssignments,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` : ``;

    const requestOptions: any = {
      url: requestUrl + propertiesSelect,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): void => {
        if (args.options.withPermissions) {
          listInstance.RoleAssignments.forEach(r => {
            r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
          });
        }

        logger.log(listInstance);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoListGetCommand();