import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  wait: boolean;
}

class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores the specified deleted site collection from tenant recycle bin';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.wait = args.options.wait;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((adminUrl: string): Promise<any> => {
        const requestOptions: any = {
          url: `${adminUrl}/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;charset=utf-8'
          },
          data: { 
            siteUrl: args.options.url
          }
        };

        return request.post(requestOptions);
      })
      .then(res => {
        logger.log(JSON.parse(res));
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '--wait'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoTenantRecycleBinItemRestoreCommand();