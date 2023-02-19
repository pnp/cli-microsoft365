import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { TenantProperty } from './TenantProperty.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl: string;
}

class SpoStorageEntityListCommand extends SpoCommand {
  public get name(): string {
    return commands.STORAGEENTITY_LIST;
  }

  public get description(): string {
    return 'Lists tenant properties stored on the specified SharePoint Online app catalog';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --appCatalogUrl <appCatalogUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.appCatalogUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving details for all tenant properties in ${args.options.appCatalogUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const web: { storageentitiesindex?: string } = await request.get<{ storageentitiesindex?: string }>(requestOptions);
      if (!web.storageentitiesindex ||
        web.storageentitiesindex.trim().length === 0) {
        if (this.verbose) {
          await logger.logToStderr('No tenant properties found');
        }
      }
      else {
        const properties: { [key: string]: TenantProperty } = JSON.parse(web.storageentitiesindex);
        const keys: string[] = Object.keys(properties);
        if (keys.length === 0) {
          if (this.verbose) {
            await logger.logToStderr('No tenant properties found');
          }
        }
        else {
          await logger.log(keys.map((key: string): any => {
            const property: TenantProperty = properties[key];
            return {
              Key: key,
              Value: property.Value,
              Description: property.Description,
              Comment: property.Comment
            };
          }));
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoStorageEntityListCommand();