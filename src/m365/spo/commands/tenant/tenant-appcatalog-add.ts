import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandError, CommandErrorWithOutput, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as spoSiteClassicAddCommand from '../site/site-classic-add';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr('Checking for existing app catalog URL...');
    }

    Cli
      .executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, { options: { _: [] } })
      .then((spoTenantAppCatalogUrlGetCommandOutput: CommandOutput): Promise<void> => {
        const appCatalogUrl: string | undefined = spoTenantAppCatalogUrlGetCommandOutput.stdout;
        if (!appCatalogUrl) {
          if (this.verbose) {
            logger.logToStderr('No app catalog URL found');
          }

          return Promise.resolve();
        }

        if (this.verbose) {
          logger.logToStderr(`Found app catalog URL ${appCatalogUrl}`);
        }

        return this.ensureNoExistingSite(appCatalogUrl, args.options.force, logger);
      })
      .then(() => this.ensureNoExistingSite(args.options.url, args.options.force, logger))
      .then(() => this.createAppCatalog(args.options, logger))
      .then(_ => cb(), (err: CommandError): void => cb(err));
  }

  private ensureNoExistingSite(url: string, force: boolean, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (err: CommandError) => void): void => {
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
      Cli
        .executeCommandWithOutput(spoSiteGetCommand as Command, siteGetOptions)
        .then(() => {
          if (this.verbose) {
            logger.logToStderr(`Found site ${url}`);
          }

          if (!force) {
            return reject(new CommandError(`Another site exists at ${url}`));
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
          }
          Cli
            .executeCommand(spoSiteRemoveCommand as Command, { options: { ...siteRemoveOptions, _: [] } })
            .then(() => resolve(), (err: CommandError) => reject(err));
        }, (err: CommandErrorWithOutput) => {
          if (err.error.message !== 'File Not Found.' && err.error.message !== '404 FILE NOT FOUND') {
            // some other error occurred
            return reject(err.error);
          }

          if (this.verbose) {
            logger.logToStderr(`No site found at ${url}`);
          }

          // site not found. continue
          resolve();
        });
    });
  }

  private createAppCatalog(options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Creating app catalog at ${options.url}...`);
    }

    const siteClassicAddOptions = {
      webTemplate: 'APPCATALOG#0',
      title: 'App catalog',
      url: options.url,
      timeZone: options.timeZone,
      owner: options.owner,
      wait: options.wait,
      verbose: this.verbose,
      debug: this.debug
    };
    return Cli.executeCommand(spoSiteClassicAddCommand as Command, { options: { ...siteClassicAddOptions, _: [] } });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (typeof args.options.timeZone !== 'number') {
      return `${args.options.timeZone} is not a number`;
    }

    return true;
  }
}

module.exports = new SpoTenantAppCatalogAddCommand();