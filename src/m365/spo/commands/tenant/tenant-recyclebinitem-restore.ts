import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { setTimeout } from 'timers/promises';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
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
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${adminUrl}/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;charset=utf-8'
        },
        data: {
          siteUrl: args.options.siteUrl
        }
      };

      const res: any = await request.post(requestOptions);
      const response = JSON.parse(res);
      if (!args.options.wait) {
        await logger.log(response);
      }
      else {
        await this.waitUntilTenantRestoreFinished(response, requestOptions, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async waitUntilTenantRestoreFinished(response: any, requestOptions: CliRequestOptions, logger: Logger): Promise<void> {
    if (response.IsComplete) {
      logger.log(response);
      return;
    }
    else {
      const pollingInterval = Number(response.PollingInterval);
      if (this.verbose) {
        logger.logToStderr(`Site collection still restoring. Retrying in ${pollingInterval / 1000} seconds...`);
      }
      await setTimeout(pollingInterval);
      const restoreResponse: any = await request.post(requestOptions);
      await this.waitUntilTenantRestoreFinished(JSON.parse(restoreResponse), requestOptions, logger);
    }
  }
}

export default new SpoTenantRecycleBinItemRestoreCommand();