import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as spoSiteAddCommand from '../site/site-add';
import * as spoSiteGetCommand from '../site/site-get';
import * as spoSiteRemoveCommand from '../site/site-remove';
import * as spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
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
      logger.logToStderr('Checking for existing app catalog URL...');
    }

    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
    const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
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

  private async ensureNoExistingSite(url: string, force: boolean, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Checking if site ${url} exists...`);
    }

    const siteGetOptions = {
      options: {
        url: url,
        verbose: this.verbose,
        debug: this.debug,
        _: []
      }
    };

    try {
      await Cli.executeCommandWithOutput(spoSiteGetCommand as Command, siteGetOptions);

      if (this.verbose) {
        logger.logToStderr(`Found site ${url}`);
      }

      if (!force) {
        throw new CommandError(`Another site exists at ${url}`);
      }

      if (this.verbose) {
        logger.logToStderr(`Deleting site ${url}...`);
      }

      const siteRemoveOptions = {
        url: url,
        skipRecycleBin: true,
        wait: true,
        confirm: true,
        verbose: this.verbose,
        debug: this.debug
      };

      await Cli.executeCommand(spoSiteRemoveCommand as Command, { options: { ...siteRemoveOptions, _: [] } });
    }
    catch (err: any) {
      if (err.message && err.message !== 'File Not Found.' && err.message !== '404 FILE NOT FOUND') {
        // Some other error occurred
        throw err.message;
      }
      else if (err.error.message !== 'File Not Found.' && err.error.message !== '404 FILE NOT FOUND') {
        // Some other error occurred
        throw err.error;
      }

      if (this.verbose) {
        logger.logToStderr(`No site found at ${url}`);
      }

      // Site not found. Continue
    }
  }

  private createAppCatalog(options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Creating app catalog at ${options.url}...`);
    }

    const siteAddOptions = {
      webTemplate: 'APPCATALOG#0',
      title: 'App catalog',
      type: 'ClassicSite',
      url: options.url,
      timeZone: options.timeZone,
      owners: options.owner,
      wait: options.wait,
      verbose: this.verbose,
      debug: this.debug,
      removeDeletedSite: false
    } as spoSiteAddCommand.Options;
    return Cli.executeCommand(spoSiteAddCommand as Command, { options: { ...siteAddOptions, _: [] } });
  }
}

module.exports = new SpoTenantAppCatalogAddCommand();