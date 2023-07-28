import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force: boolean;
  owner: string;
  timeZone: string | number;
  url: string;
  wait: boolean;
}

class SpoTenantAppCatalogAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOG_ADD;
  }

  public get description(): string {
    return 'Creates new tenant app catalog site';
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
        wait: args.options.wait || false,
        force: args.options.force || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--owner <owner>'
      },
      {
        option: '-z, --timeZone <timeZone>'
      },
      {
        option: '--wait'
      },
      {
        option: '--force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (typeof args.options.timeZone !== 'number') {
          return `${args.options.timeZone} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Checking for existing app catalog URL...');
    }
    try {
      const appCatalogUrl: string | null = await spo.getTenantAppCatalogUrl(logger, this.verbose);
      if (!appCatalogUrl) {
        if (this.verbose) {
          logger.logToStderr('No app catalog URL found');
        }
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Found app catalog URL ${appCatalogUrl}`);
        }

        //Using JSON.parse
        await this.ensureNoExistingSite(appCatalogUrl, args.options.force, logger);
      }
      await this.ensureNoExistingSite(args.options.url, args.options.force, logger);
      await this.createAppCatalog(args.options, logger);
    }
    catch (ex) {
      this.handleRejectedODataPromise(ex);
    }
  }

  private async ensureNoExistingSite(url: string, force: boolean, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if site ${url} exists...`);
    }

    try {
      await spo.getSite(url, logger, this.verbose);

      if (this.verbose) {
        await logger.logToStderr(`Found site ${url}`);
      }

      if (!force) {
        throw new CommandError(`Another site exists at ${url}`);
      }

      if (this.verbose) {
        await logger.logToStderr(`Deleting site ${url}...`);
      }

      await spo.removeSite(url, true, true, logger, this.verbose);
    }
    catch (err: any) {
      logger.log(err);
      if (err.message !== 'File not Found' && err.message !== '404 FILE NOT FOUND') {
        throw err.message;
      }

      if (this.verbose) {
        await logger.logToStderr(`No site found at ${url}`);
      }

      // Site not found. Continue
    }
  }

  private async createAppCatalog(options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating app catalog at ${options.url}...`);
    }

    return await spo.addSite(
      'App catalog',
      logger,
      this.verbose,
      options.wait,
      'ClassicSite',
      undefined,
      undefined,
      options.owner,
      undefined,
      false,
      undefined,
      undefined,
      undefined,
      options.url,
      undefined,
      undefined,
      options.timeZone,
      'APPCATALOG#0');
  }
}

export default new SpoTenantAppCatalogAddCommand();