import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { FormDigestInfo, spo } from '../../../../utils/spo.js';
import { timersUtil } from '../../../../utils/timersUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  newUrl: string;
  newTitle?: string;
  suppressMarketplaceAppCheck?: boolean;
  suppressWorkflow2013Check?: boolean;
  wait?: boolean;
}

interface SiteRenameJob {
  ErrorDescription: string;
  JobState: string;
}

class SpoTenantSiteRenameCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private operationData?: SiteRenameJob;
  private static readonly checkIntervalInMs: number = 5000;

  public get name(): string {
    return commands.TENANT_SITE_RENAME;
  }

  public get description(): string {
    return 'Renames the URL and title of a site collection';
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
        newTitle: typeof args.options.newTitle !== 'undefined',
        suppressMarketplaceAppCheck: !!args.options.suppressMarketplaceAppCheck,
        suppressWorkflow2013Check: !!args.options.suppressWorkflow2013Check,
        wait: !!args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--newUrl <newUrl>'
      },
      {
        option: '--newTitle [newTitle]'
      },
      {
        option: '--suppressMarketplaceAppCheck'
      },
      {
        option: '--suppressWorkflow2013Check'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.url.toLowerCase() === args.options.newUrl.toLowerCase()) {
          return 'The new URL cannot be the same as the target URL.';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const options = args.options;
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      const reqDigest = await spo.getRequestDigest(spoAdminUrl);
      this.context = reqDigest;
      if (this.verbose) {
        await logger.logToStderr(`Scheduling rename job...`);
      }

      let optionsBitmask = 0;
      if (options.suppressMarketplaceAppCheck) {
        optionsBitmask = optionsBitmask | 8;
      }

      if (options.suppressWorkflow2013Check) {
        optionsBitmask = optionsBitmask | 16;
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SiteRenameJobs?api-version=1.4.7`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue,
          'Content-Type': 'application/json'
        },
        responseType: 'json',
        data: {
          SourceSiteUrl: options.url,
          TargetSiteUrl: options.newUrl,
          TargetSiteTitle: options.newTitle || null,
          Option: optionsBitmask,
          Reserve: null,
          SkipGestures: null,
          OperationId: '00000000-0000-0000-0000-000000000000'
        }
      };

      const res = await request.post<SiteRenameJob>(requestOptions);

      if (options.verbose) {
        await logger.logToStderr(res);
      }

      this.operationData = res;

      if (this.operationData.JobState && this.operationData.JobState === "Error") {
        throw this.operationData.ErrorDescription;
      }

      const isComplete: boolean = this.operationData.JobState === "Success";
      if (options.wait && !isComplete) {
        await this.waitForRenameCompletion(
          this,
          true,
          spoAdminUrl,
          options.url,
          0
        );
      }
      await logger.log(this.operationData);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  protected async waitForRenameCompletion(command: SpoTenantSiteRenameCommand, isVerbose: boolean, spoAdminUrl: string, siteUrl: string, iteration: number): Promise<void> {
    iteration++;

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SiteRenameJobs/GetJobsBySiteUrl(url='${formatting.encodeQueryParameter(siteUrl)}')?api-version=1.4.7`,
      headers: {
        'X-AttemptNumber': iteration.toString()
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: SiteRenameJob[] }>(requestOptions);
    this.operationData = res.value[0];

    if (this.operationData.ErrorDescription) {
      throw this.operationData.ErrorDescription;
    }

    if (this.operationData.JobState === "Success") {
      return;
    }

    await timersUtil.setTimeout(SpoTenantSiteRenameCommand.checkIntervalInMs);
    await command.waitForRenameCompletion(command, isVerbose, spoAdminUrl, siteUrl, iteration);
  }
}

export default new SpoTenantSiteRenameCommand();