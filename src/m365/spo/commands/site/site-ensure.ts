import * as chalk from 'chalk';
import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, {
  CommandOption,
  CommandErrorWithOutput
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as spoWebGetCommand from '../web/web-get';
import { Options as SpoWebGetCommandOptions } from '../web/web-get';
import { SharingCapabilities } from './SharingCapabilities';
import * as spoSiteAddCommand from './site-add';
import { Options as SpoSiteAddCommandOptions } from './site-add';
import * as spoSiteSetCommand from './site-set';
import { Options as SpoSiteSetCommandOptions } from './site-set';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  // add
  type?: string;
  title: string;
  alias?: string;
  description?: string;
  classification?: string;
  isPublic?: boolean;
  lcid?: number;
  owners: string;
  shareByEmailEnabled?: boolean;
  siteDesign?: string;
  siteDesignId?: string;
  timeZone?: string | number;
  webTemplate?: string;
  resourceQuota?: string | number;
  resourceQuotaWarningLevel?: string | number;
  storageQuota?: string | number;
  storageQuotaWarningLevel?: string | number;
  removeDeletedSite: boolean;
  wait: boolean;
  // set
  disableFlows?: string;
  sharingCapability?: string;
}

class SpoSiteEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ENSURE;
  }

  public get description(): string {
    return 'Ensures that the particular site collection exists and updates its properties if necessary';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getWeb(args, logger)
      .then((getWebOutput: CommandOutput): Promise<CommandOutput> => {
        if (this.debug) {
          logger.logToStderr(getWebOutput.stderr);
        }

        if (this.verbose) {
          logger.logToStderr(`Site found at ${args.options.url}. Checking if site matches conditions...`);
        }

        const web: {
          Configuration: number;
          WebTemplate: string;
        } = JSON.parse(getWebOutput.stdout);

        if (args.options.type) {
          // type was specified so we need to check if the existing site matches
          // it. If not, we throw an error and stop
          // Determine the type of site to match
          let expectedWebTemplate: string | undefined;
          switch (args.options.type) {
            case 'TeamSite':
              expectedWebTemplate = 'GROUP#0';
              break;
            case 'CommunicationSite':
              expectedWebTemplate = 'SITEPAGEPUBLISHING#0';
              break;
            case 'ClassicSite':
              expectedWebTemplate = args.options.webTemplate;
              break;
            default:
              return Promise.reject(`${args.options.type} is not a valid site type. Allowed types are TeamSite,CommunicationSite,ClassicSite`);
          }

          if (expectedWebTemplate) {
            const currentWebTemplate = `${web.WebTemplate}#${web.Configuration}`;
            if (expectedWebTemplate !== currentWebTemplate) {
              return Promise.reject(`Expected web template ${expectedWebTemplate} but site found at ${args.options.url} is based on ${currentWebTemplate}`);
            }
          }
        }

        if (this.verbose) {
          logger.logToStderr(`Site matches conditions. Updating...`);
        }

        return this.updateSite(args, logger);
      }, (err: CommandErrorWithOutput): Promise<CommandOutput> => {
        if (this.debug) {
          logger.logToStderr(err.stderr);
        }

        if (err.error.message !== 'Request failed with status code 404') {
          return Promise.reject(err);
        }

        if (this.verbose) {
          logger.logToStderr(`No site found at ${args.options.url}`);
        }

        return this.createSite(args, logger);
      })
      .then((res: CommandOutput): void => {
        if (this.debug) {
          logger.logToStderr(res.stderr);
        }

        logger.log(res.stdout);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getWeb(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Checking if site ${args.options.url} exists...`);
    }

    const options: SpoWebGetCommandOptions = {
      webUrl: args.options.url,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };
    return Cli.executeCommandWithOutput(spoWebGetCommand as Command, { options: { ...options, _: [] } });
  }

  private createSite(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Creating site...`);
    }

    const options: SpoSiteAddCommandOptions = {
      type: args.options.type,
      title: args.options.title,
      alias: args.options.alias,
      description: args.options.description,
      classification: args.options.classification,
      isPublic: args.options.isPublic,
      lcid: args.options.lcid,
      url: typeof args.options.type === 'undefined' || args.options.type === 'TeamSite' ? undefined : args.options.url,
      owners: args.options.owners,
      shareByEmailEnabled: args.options.shareByEmailEnabled,
      siteDesign: args.options.siteDesign,
      siteDesignId: args.options.siteDesignId,
      timeZone: args.options.timeZone,
      webTemplate: args.options.webTemplate,
      resourceQuota: args.options.resourceQuota,
      resourceQuotaWarningLevel: args.options.resourceQuotaWarningLevel,
      storageQuota: args.options.storageQuota,
      storageQuotaWarningLevel: args.options.storageQuotaWarningLevel,
      removeDeletedSite: args.options.removeDeletedSite,
      wait: args.options.wait,
      verbose: this.verbose,
      debug: this.debug
    };

    const validationResult: boolean | string = (spoSiteAddCommand as Command).validate({ options: options });
    if (validationResult !== true) {
      return Promise.reject(validationResult);
    }

    return Cli.executeCommandWithOutput(spoSiteAddCommand as Command, { options: { ...options, _: [] } });
  }

  private updateSite(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Updating site...`);
    }

    const validationResult: boolean | string = (spoSiteSetCommand as Command).validate(args);
    if (validationResult !== true) {
      return Promise.reject(validationResult);
    }

    const options: SpoSiteSetCommandOptions = {
      classification: args.options.classification,
      disableFlows: args.options.disableFlows,
      isPublic: typeof args.options.isPublic !== 'undefined' ? args.options.isPublic.toString() : undefined,
      owners: args.options.owners,
      shareByEmailEnabled: typeof args.options.shareByEmailEnabled !== 'undefined' ? args.options.shareByEmailEnabled.toString() : undefined,
      siteDesignId: args.options.siteDesignId,
      title: args.options.title,
      url: args.options.url,
      sharingCapability: args.options.sharingCapability,
      verbose: this.verbose,
      debug: this.debug
    };
    return Cli.executeCommandWithOutput(spoSiteSetCommand as Command, { options: { ...options, _: [] } });
  }

  /**
   * Maps the base sharingCapability enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  private get sharingCapabilities(): string[] {
    const result: string[] = [];

    for (const sharingCapability in SharingCapabilities) {
      if (typeof SharingCapabilities[sharingCapability] === 'number') {
        result.push(sharingCapability);
      }
    }

    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '--type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite', 'ClassicSite']
      },
      {
        option: '-t, --title <title>'
      },
      {
        option: '-a, --alias [alias]'
      },
      {
        option: '-z, --timeZone [timeZone]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-l, --lcid [lcid]'
      },
      {
        option: '--owners [owners]'
      },
      {
        option: '--isPublic'
      },
      {
        option: '-c, --classification [classification]'
      },
      {
        option: '--siteDesign [siteDesign]',
        autocomplete: ['Topic', 'Showcase', 'Blank']
      },
      {
        option: '--siteDesignId [siteDesignId]'
      },
      {
        option: '--shareByEmailEnabled'
      },
      {
        option: '-w, --webTemplate [webTemplate]'
      },
      {
        option: '--resourceQuota [resourceQuota]'
      },
      {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]'
      },
      {
        option: '--storageQuota [storageQuota]'
      },
      {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]'
      },
      {
        option: '--removeDeletedSite'
      },
      {
        option: '--disableFlows [disableFlows]'
      },
      {
        option: '--sharingCapability [sharingCapability]',
        autocomplete: this.sharingCapabilities
      },
      {
        option: '--wait'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoSiteEnsureCommand();