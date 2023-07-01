import { Cli, CommandOutput } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { CommandError } from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoSiteAddCommand, { Options as SpoSiteAddCommandOptions } from '../site/site-add.js';
import spoSiteGetCommand from '../site/site-get.js';
import spoSiteRemoveCommand from '../site/site-remove.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';

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
      await logger.logToStderr('Checking for existing app catalog URL...');
    }

    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await Cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
    const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
    if (!appCatalogUrl) {
      if (this.verbose) {
        await logger.logToStderr('No app catalog URL found');
      }
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Found app catalog URL ${appCatalogUrl}`);
      }

      //Using JSON.parse
      await this.ensureNoExistingSite(appCatalogUrl, args.options.force, logger);
    }
    await this.ensureNoExistingSite(args.options.url, args.options.force, logger);
    await this.createAppCatalog(args.options, logger);
  }

  private async ensureNoExistingSite(url: string, force: boolean, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if site ${url} exists...`);
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
        await logger.logToStderr(`Found site ${url}`);
      }

      if (!force) {
        throw new CommandError(`Another site exists at ${url}`);
      }

      if (this.verbose) {
        await logger.logToStderr(`Deleting site ${url}...`);
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
      if (err.error?.message !== 'File not Found' && err.error?.message !== '404 FILE NOT FOUND') {
        throw err.error || err;
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
    } as SpoSiteAddCommandOptions;
    return Cli.executeCommand(spoSiteAddCommand as Command, { options: { ...siteAddOptions, _: [] } });
  }
}

export default new SpoTenantAppCatalogAddCommand();