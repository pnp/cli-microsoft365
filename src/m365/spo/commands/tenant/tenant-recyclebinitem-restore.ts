import { Logger } from '../../../../cli';
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        wait: args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
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
}

module.exports = new SpoTenantRecycleBinItemRestoreCommand();