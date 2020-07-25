import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { TenantProperty } from './TenantProperty';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(cmd, this.debug)
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
            cmd.log(`Property with key ${args.options.key} not found`);
          }
        }
        else {
          cmd.log({
            Key: args.options.key,
            Value: property.Value,
            Description: property.Description,
            Comment: property.Comment
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
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