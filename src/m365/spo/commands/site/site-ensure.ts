import * as chalk from 'chalk';
import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
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
  disableFlows?: boolean;
  sharingCapability?: string;
}

class SpoSiteEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ENSURE;
  }

  public get description(): string {
    return 'Ensures that the particular site collection exists and updates its properties if necessary';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '--disableFlows [disableFlows]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--sharingCapability [sharingCapability]',
        autocomplete: this.sharingCapabilities
      },
      {
        option: '--wait'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('disableFlows');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res = await this.ensureSite(logger, args);

      if (this.debug) {
        logger.logToStderr(res.stderr);
      }

      logger.log(res.stdout);

      if (this.verbose) {
        logger.logToStderr(chalk.green('DONE'));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureSite(logger: Logger, args: CommandArgs): Promise<CommandOutput> {
    let getWebOutput: CommandOutput;
    try {
      getWebOutput = await this.getWeb(args, logger);
    }
    catch (err: any) {
      if (this.debug) {
        logger.logToStderr(err.stderr);
      }

      if (err.error.message !== '404 FILE NOT FOUND') {
        throw err;
      }

      if (this.verbose) {
        logger.logToStderr(`No site found at ${args.options.url}`);
      }

      return this.createSite(args, logger);
    }

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
          throw `${args.options.type} is not a valid site type. Allowed types are TeamSite,CommunicationSite,ClassicSite`;
      }

      if (expectedWebTemplate) {
        const currentWebTemplate = `${web.WebTemplate}#${web.Configuration}`;
        if (expectedWebTemplate !== currentWebTemplate) {
          throw `Expected web template ${expectedWebTemplate} but site found at ${args.options.url} is based on ${currentWebTemplate}`;
        }
      }
    }

    if (this.verbose) {
      logger.logToStderr(`Site matches conditions. Updating...`);
    }

    return this.updateSite(args, logger);
  }

  private getWeb(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Checking if site ${args.options.url} exists...`);
    }

    const options: SpoWebGetCommandOptions = {
      url: args.options.url,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };
    return Cli.executeCommandWithOutput(spoWebGetCommand as Command, { options: { ...options, _: [] } });
  }

  private async createSite(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
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

    const validationResult: boolean | string = await (spoSiteAddCommand as Command).validate({ options: options }, Cli.getCommandInfo(spoSiteAddCommand as Command));
    if (validationResult !== true) {
      return Promise.reject(validationResult);
    }

    return Cli.executeCommandWithOutput(spoSiteAddCommand as Command, { options: { ...options, _: [] } });
  }

  private async updateSite(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Updating site...`);
    }

    const options: SpoSiteSetCommandOptions = {
      classification: args.options.classification,
      disableFlows: args.options.disableFlows,
      isPublic: args.options.isPublic,
      owners: args.options.owners,
      shareByEmailEnabled: args.options.shareByEmailEnabled,
      siteDesignId: args.options.siteDesignId,
      title: args.options.title,
      url: args.options.url,
      sharingCapability: args.options.sharingCapability,
      verbose: this.verbose,
      debug: this.debug
    };
    const validationResult: boolean | string = await (spoSiteSetCommand as Command).validate({ options: options }, Cli.getCommandInfo(spoSiteSetCommand as Command));
    if (validationResult !== true) {
      return Promise.reject(validationResult);
    }

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
}

module.exports = new SpoSiteEnsureCommand();