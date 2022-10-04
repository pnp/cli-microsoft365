import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { TenantProperty } from './TenantProperty';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key: string;
}

class SpoStorageEntityGetCommand extends SpoCommand {
  public get name(): string {
    return commands.STORAGEENTITY_GET;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
  }

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-k, --key <key>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/web/GetStorageEntity('${encodeURIComponent(args.options.key)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const property: TenantProperty = await request.get(requestOptions);
      if (property["odata.null"] === true) {
        if (this.verbose) {
          logger.logToStderr(`Property with key ${args.options.key} not found`);
        }
      }
      else {
        logger.log({
          Key: args.options.key,
          Value: property.Value,
          Description: property.Description,
          Comment: property.Comment
        });
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoStorageEntityGetCommand();