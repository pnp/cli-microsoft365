import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListPrincipalType } from '../list/ListPrincipalType';
import { FolderProperties } from './FolderProperties';
interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  url?: string;
  id?: string;
  withPermissions?: boolean;
}

class SpoFolderGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified folder';
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
        url: typeof args.options.url !== 'undefined',
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
        option: '-f, --url [url]'
      },
      {
        option: '-i, --id [id]'
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
    this.optionSets.push({ options: ['url', 'id'] });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving folder from site ${args.options.webUrl}...`);
    }
    let requestUrl: string = `${args.options.webUrl}/_api/web`;
    if (args.options.id) {
      requestUrl += `/GetFolderById('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.url) {
      const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl += `/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativePath)}')`;
    }
    if (args.options.withPermissions) {
      requestUrl += `?$expand=ListItemAllFields/HasUniqueRoleAssignments,ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const folder = await request.get<FolderProperties>(requestOptions);
      if (args.options.withPermissions) {
        const listItemAllFields = folder.ListItemAllFields;
        if (!(listItemAllFields ?? false)) {
          throw Error('Please ensure the specified folder URL or folder Id does not refer to a root folder. Use \'spo list get\' with withPermissions instead.');
        }
        listItemAllFields.RoleAssignments.forEach(r => {
          r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
      }
      logger.log(folder);
    }
    catch (err: any) {
      if (err.statusCode && err.statusCode === 500) {
        throw new CommandError('Please check the folder URL. Folder might not exist on the specified URL');
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFolderGetCommand();