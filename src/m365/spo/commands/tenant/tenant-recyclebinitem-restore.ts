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

interface Response {
  HasTimedout: boolean,
  IsComplete: boolean,
  PollingInterval: number
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

      const response: string = await request.post(requestOptions);
      let responseContent: Response = JSON.parse(response);

      if (args.options.wait && !responseContent.IsComplete) {
        responseContent = await this.waitUntilTenantRestoreFinished(responseContent.PollingInterval, requestOptions, logger);
      }

      await logger.log(responseContent);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async waitUntilTenantRestoreFinished(pollingInterval: number, requestOptions: CliRequestOptions, logger: Logger): Promise<any> {
    if (this.verbose) {
      await logger.logToStderr(`Site collection still restoring. Retrying in ${pollingInterval / 1000} seconds...`);
    }

    await setTimeout(pollingInterval);

    const response: string = await request.post(requestOptions);
    const responseContent: Response = JSON.parse(response);

    if (responseContent.IsComplete) {
      return responseContent;
    }

    return await this.waitUntilTenantRestoreFinished(responseContent.PollingInterval, requestOptions, logger);
  }
}

export default new SpoTenantRecycleBinItemRestoreCommand();