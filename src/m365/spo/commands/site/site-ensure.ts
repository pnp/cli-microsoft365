import * as chalk from 'chalk';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SharingCapabilities } from './SharingCapabilities';
import { spo } from '../../../../utils/spo';
import { WebProperties } from '../web/WebProperties';

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

      logger.log(res);

      if (this.verbose) {
        logger.logToStderr(chalk.green('DONE'));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureSite(logger: Logger, args: CommandArgs): Promise<any> {
    let getWebOutput: WebProperties;
    try {
      getWebOutput = await this.getWeb(args, logger);
    }
    catch (err: any) {
      if (this.debug) {
        logger.logToStderr(err);
      }

      if (err.message !== '404 FILE NOT FOUND') {
        throw err;
      }

      if (this.verbose) {
        logger.logToStderr(`No site found at ${args.options.url}`);
      }

      return this.createSite(args, logger);
    }

    if (this.debug) {
      logger.logToStderr(getWebOutput);
    }

    if (this.verbose) {
      logger.logToStderr(`Site found at ${args.options.url}. Checking if site matches conditions...`);
    }

    const web: {
      Configuration: number;
      WebTemplate: string;
    } = {
      Configuration: getWebOutput.Configuration,
      WebTemplate: getWebOutput.WebTemplate
    };

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
          expectedWebTemplate = args.options.webTemplate || 'STS#0';
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

  private async getWeb(args: CommandArgs, logger: Logger): Promise<WebProperties> {
    if (this.verbose) {
      logger.logToStderr(`Checking if site ${args.options.url} exists...`);
    }

    return await spo.getWeb(args.options.url, logger, this.verbose);
  }

  private async createSite(args: CommandArgs, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Creating site...`);
    }

    const url = typeof args.options.type === 'undefined' || args.options.type === 'TeamSite' ? undefined : args.options.url;

    return await spo.addSite(
      args.options.title,
      logger,
      this.verbose,
      args.options.wait,
      args.options.type,
      args.options.alias,
      args.options.description,
      args.options.owners,
      args.options.shareByEmailEnabled,
      args.options.removeDeletedSite,
      args.options.classification,
      args.options.isPublic,
      args.options.lcid,
      url,
      args.options.siteDesign,
      args.options.siteDesignId,
      args.options.timeZone,
      args.options.webTemplate,
      args.options.resourceQuota,
      args.options.resourceQuotaWarningLevel,
      args.options.storageQuota,
      args.options.storageQuotaWarningLevel
    );
  }

  private async updateSite(args: CommandArgs, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Updating site...`);
    }

    return await spo.updateSite(
      args.options.url,
      logger,
      this.verbose,
      args.options.title,
      args.options.classification,
      args.options.disableFlows,
      args.options.isPublic,
      args.options.owners,
      args.options.shareByEmailEnabled,
      args.options.siteDesignId,
      args.options.sharingCapability
    );
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