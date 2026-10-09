import { z } from 'zod';
import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { CommandError, globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoSiteAddCommand, { Options as SpoSiteAddCommandOptions } from '../site/site-add.js';
import spoSiteGetCommand from '../site/site-get.js';
import spoSiteRemoveCommand from '../site/site-remove.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string().refine(val => validation.isValidSharePointUrl(val) === true, { message: 'The value is not a valid SharePoint site URL.' }).alias('u'),
  owner: z.string().optional(),
  timeZone: z.string().refine(val => !isNaN(Number(val)), { message: 'timeZone is not a number' }).optional().alias('z'),
  wait: z.boolean().optional(),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantAppCatalogAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOG_ADD;
  }

  public get description(): string {
    return 'Creates new tenant app catalog site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Checking for existing app catalog URL...');
    }

    const spoTenantAppCatalogUrlGetCommandOutput: CommandOutput = await cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { output: 'text', _: [] } });
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
      await this.ensureNoExistingSite(appCatalogUrl, args.options.force ?? false, logger);
    }
    await this.ensureNoExistingSite(args.options.url, args.options.force ?? false, logger);
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
      await cli.executeCommandWithOutput(spoSiteGetCommand as Command, siteGetOptions);

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
        permanent: true,
        wait: true,
        force: true,
        verbose: this.verbose,
        debug: this.debug
      };

      await cli.executeCommand(spoSiteRemoveCommand as Command, { options: { ...siteRemoveOptions, _: [] } });
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
    return cli.executeCommand(spoSiteAddCommand as Command, { options: { ...siteAddOptions, _: [] } });
  }
}

export default new SpoTenantAppCatalogAddCommand();