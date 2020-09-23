import { Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
    return `${commands.STORAGEENTITY_GET}`;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<TenantProperty> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/web/GetStorageEntity('${encodeURIComponent(args.options.key)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((property: TenantProperty): void => {
        if (property["odata.null"] === true) {
          if (this.verbose) {
            logger.log(`Property with key ${args.options.key} not found`);
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
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-k, --key <key>',
      description: 'Name of the tenant property to retrieve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoStorageEntityGetCommand();