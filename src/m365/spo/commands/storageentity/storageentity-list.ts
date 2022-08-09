import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { TenantProperty } from './TenantProperty';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving details for all tenant properties in ${args.options.appCatalogUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ storageentitiesindex?: string }>(requestOptions)
      .then((web: { storageentitiesindex?: string }): void => {
        try {
          if (!web.storageentitiesindex ||
            web.storageentitiesindex.trim().length === 0) {
            if (this.verbose) {
              logger.logToStderr('No tenant properties found');
            }
            cb();
            return;
          }

          const properties: { [key: string]: TenantProperty } = JSON.parse(web.storageentitiesindex);
          const keys: string[] = Object.keys(properties);
          if (keys.length === 0) {
            if (this.verbose) {
              logger.logToStderr('No tenant properties found');
            }
          }
          else {
            logger.log(keys.map((key: string): any => {
              const property: TenantProperty = properties[key];
              return {
                Key: key,
                Value: property.Value,
                Description: property.Description,
                Comment: property.Comment
              };
            }));
          }
          cb();
        }
        catch (e) {
          this.handleError(e, logger, cb);
        }
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoStorageEntityListCommand();