import * as chalk from 'chalk';
import { Cli, CommandErrorWithOutput, CommandOutput, Logger } from '../../../../cli';
import Command, {
  CommandOption
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

    for (let sharingCapability in SharingCapabilities) {
      if (typeof SharingCapabilities[sharingCapability] === 'number') {
        result.push(sharingCapability);
      }
    }

    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site collection to ensure that it exists and is properly configured'
      },
      {
        option: '--type [type]',
        description: 'Type of sites to add. Allowed values TeamSite|CommunicationSite|ClassicSite, default TeamSite',
        autocomplete: ['TeamSite', 'CommunicationSite', 'ClassicSite']
      },
      {
        option: '-t, --title <title>',
        description: 'Site title'
      },
      {
        option: '-a, --alias [alias]',
        description: 'Site alias, used in the URL and in the team site group e-mail (applies to type TeamSite)'
      },
      {
        option: '-z, --timeZone [timeZone]',
        description: 'Integer representing time zone to use for the site (applies to type ClassicSite)'
      },
      {
        option: '-d, --description [description]',
        description: 'Site description'
      },
      {
        option: '-l, --lcid [lcid]',
        description: 'Site language in the LCID format, eg. 1033 for en-US. See https://support.microsoft.com/en-us/office/languages-supported-by-sharepoint-dfbf3652-2902-4809-be21-9080b6512fff for the list of supported languages'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of users to set as site owners'
      },
      {
        option: '--isPublic',
        description: 'Determines if the associated group is public or not (applies to type TeamSite)'
      },
      {
        option: '-c, --classification [classification]',
        description: 'Site classification (applies to type TeamSite, CommunicationSite)'
      },
      {
        option: '--siteDesign [siteDesign]',
        description: 'Type of communication site to create. Allowed values Topic|Showcase|Blank, default Topic. Specify either siteDesign or siteDesignId (applies to type CommunicationSite)',
        autocomplete: ['Topic', 'Showcase', 'Blank']
      },
      {
        option: '--siteDesignId [siteDesignId]',
        description: 'Id of the custom site design to use to create the site. Specify either siteDesign or siteDesignId (applies to type CommunicationSite)'
      },
      {
        option: '--shareByEmailEnabled',
        description: 'Determines whether it\'s allowed to share file with guests (applies to type CommunicationSite)'
      },
      {
        option: '-w, --webTemplate [webTemplate]',
        description: 'Template to use for creating the site. Default STS#0 (applies to type ClassicSite)'
      },
      {
        option: '--resourceQuota [resourceQuota]',
        description: 'The quota for this site collection in Sandboxed Solutions units. Default 0 (applies to type ClassicSite)'
      },
      {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]',
        description: 'The warning level for the resource quota. Default 0 (applies to type ClassicSite)'
      },
      {
        option: '--storageQuota [storageQuota]',
        description: 'The storage quota for this site collection in megabytes. Default 100 (applies to type ClassicSite)'
      },
      {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]',
        description: 'The warning level for the storage quota in megabytes. Default 100 (applies to type ClassicSite)'
      },
      {
        option: '--removeDeletedSite',
        description: 'Set, to remove existing deleted site with the same URL from the Recycle Bin (applies to type ClassicSite)'
      },
      {
        option: '--disableFlows [disableFlows]',
        description: 'Set to true to disable using Microsoft Flow in this site collection'
      },
      {
        option: '--sharingCapability [sharingCapability]',
        description: `The sharing capability for the Site. Allowed values ${this.sharingCapabilities.join('|')}.`,
        autocomplete: this.sharingCapabilities
      },
      {
        option: '--wait',
        description: 'Wait for the site to be provisioned before completing the command (applies to type ClassicSite)'
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