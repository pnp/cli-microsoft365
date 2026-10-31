import chalk from 'chalk';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { WebProperties } from '../web/WebProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  type: z.enum(['TeamSite', 'CommunicationSite', 'ClassicSite']).optional(),
  title: z.string().alias('t'),
  alias: z.string().optional().alias('a'),
  timeZone: z.number().optional().alias('z'),
  description: z.string().optional().alias('d'),
  lcid: z.number().optional().alias('l'),
  owners: z.string().optional(),
  isPublic: z.boolean().optional(),
  classification: z.string().optional().alias('c'),
  siteDesign: z.enum(['Topic', 'Showcase', 'Blank']).optional(),
  siteDesignId: z.string().optional(),
  shareByEmailEnabled: z.boolean().optional(),
  webTemplate: z.string().optional().alias('w'),
  resourceQuota: z.number().optional(),
  resourceQuotaWarningLevel: z.number().optional(),
  storageQuota: z.number().optional(),
  storageQuotaWarningLevel: z.number().optional(),
  removeDeletedSite: z.boolean().optional(),
  disableFlows: z.boolean().optional(),
  sharingCapability: z.enum(['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']).optional(),
  wait: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ENSURE;
  }

  public get description(): string {
    return 'Ensures that the particular site collection exists and updates its properties if necessary';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res = await this.ensureSite(logger, args);

      await logger.log(res);

      if (this.verbose) {
        await logger.logToStderr(chalk.green('DONE'));
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
        await logger.logToStderr(err);
      }

      if (err.error !== '404 FILE NOT FOUND') {
        throw err;
      }

      if (this.verbose) {
        await logger.logToStderr(`No site found at ${args.options.url}`);
      }

      return this.createSite(args, logger);
    }

    if (this.debug) {
      await logger.logToStderr(getWebOutput);
    }

    if (this.verbose) {
      await logger.logToStderr(`Site found at ${args.options.url}. Checking if site matches conditions...`);
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
      }

      if (expectedWebTemplate) {
        const currentWebTemplate = `${web.WebTemplate}#${web.Configuration}`;
        if (expectedWebTemplate !== currentWebTemplate) {
          throw `Expected web template ${expectedWebTemplate} but site found at ${args.options.url} is based on ${currentWebTemplate}`;
        }
      }
    }

    if (this.verbose) {
      await logger.logToStderr(`Site matches conditions. Updating...`);
    }

    return this.updateSite(args, logger);
  }

  private async getWeb(args: CommandArgs, logger: Logger): Promise<WebProperties> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if site ${args.options.url} exists...`);
    }

    return await spo.getWeb(args.options.url, logger, this.verbose);
  }

  private async createSite(args: CommandArgs, logger: Logger): Promise<any> {
    if (this.verbose) {
      await logger.logToStderr(`Creating site...`);
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
      await logger.logToStderr(`Updating site...`);
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
}

export default new SpoSiteEnsureCommand();