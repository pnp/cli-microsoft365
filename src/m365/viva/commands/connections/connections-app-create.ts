import AdmZip from 'adm-zip';
import fs from 'fs';
import path from 'path';
import { v4 } from 'uuid';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
import { WebProperties } from '../../../spo/commands/web/WebProperties.js';
import { spo } from '../../../../utils/spo.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  portalUrl: z.string(),
  name: z.string().max(30, { message: 'App name must not exceed 30 characters' }),
  description: z.string().max(80, { message: 'Description must not exceed 80 characters' }),
  longDescription: z.string().max(4000, { message: 'Long description must not exceed 4000 characters' }),
  privacyPolicyUrl: z.string().optional(),
  termsOfUseUrl: z.string().optional(),
  companyName: z.string(),
  companyWebsiteUrl: z.string(),
  coloredIconPath: z.string().refine(val => fs.existsSync(path.resolve(val)), {
    error: e => `File ${path.resolve(e.input as string)} doesn't exist`
  }),
  outlineIconPath: z.string().refine(val => fs.existsSync(path.resolve(val)), {
    error: e => `File ${path.resolve(e.input as string)} doesn't exist`
  }),
  accentColor: z.string().optional(),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaConnectionsAppCreateCommand extends AnonymousCommand {
  private archive?: AdmZip;

  public get name(): string {
    return commands.CONNECTIONS_APP_CREATE;
  }

  public get description(): string {
    return 'Creates Viva Connections app';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .superRefine((opts, ctx) => {
        const appFilePath = path.resolve(`${opts.name}.zip`);
        if (fs.existsSync(appFilePath) && !opts.force) {
          ctx.addIssue({
            message: `File ${path.resolve(`${opts.name}.zip`)} already exists. Delete the file or use the --force option to overwrite the existing file`,
            code: z.ZodIssueCode.custom,
            params: { customCode: 'required' }
          });
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const web: WebProperties = await this.getWeb(args, logger);
      if (this.debug) {
        await logger.logToStderr(web);
      }

      if (this.verbose) {
        await logger.logToStderr(`Site found at ${args.options.portalUrl}. Checking if it's a communication site...`);
      }

      if (web.WebTemplate !== 'SITEPAGEPUBLISHING' ||
        web.Configuration !== 0) {
        throw `Site ${args.options.portalUrl} is not a Communication Site. Please specify a different site and try again.`;
      }

      if (this.verbose) {
        await logger.logToStderr(`Site ${args.options.portalUrl} is a Communication Site. Building app...`);
      }

      const portalUrl: URL = new URL(args.options.portalUrl);
      const appPortalUrl: string = `${args.options.portalUrl}${args.options.portalUrl.indexOf('?') > -1 ? '&' : '?'}app=portals`;
      let searchUrlPath: string = portalUrl.hostname;
      if (portalUrl.pathname.indexOf('/teams') > -1 || portalUrl.pathname.indexOf('/sites') > -1) {
        const firstTwoUrlSegments = portalUrl.pathname.match(/^\/[^/]+\/[^/]+/);
        if (firstTwoUrlSegments) {
          searchUrlPath += firstTwoUrlSegments[0];
        }
      }
      const coloredIconPath = path.resolve(args.options.coloredIconPath);
      const coloredIconFileName: string = path.basename(coloredIconPath);
      const outlineIconPath = path.resolve(args.options.outlineIconPath);
      const outlineIconFileName: string = path.basename(outlineIconPath);
      const domain: string = portalUrl.hostname;
      const appId: string = v4();

      const manifest: any = {
        "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.9/MicrosoftTeams.schema.json",
        "manifestVersion": "1.9",
        "version": "1.0",
        "id": appId,
        "packageName": `com.microsoft.teams.${args.options.name}`,
        "developer": {
          "name": args.options.companyName,
          "websiteUrl": args.options.companyWebsiteUrl,
          "privacyUrl": args.options.privacyPolicyUrl || 'https://privacy.microsoft.com/en-us/privacystatement',
          "termsOfUseUrl": args.options.termsOfUseUrl || 'https://go.microsoft.com/fwlink/?linkid=2039674'
        },
        "icons": {
          "color": coloredIconFileName,
          "outline": outlineIconFileName
        },
        "name": {
          "short": args.options.name,
          "full": args.options.name
        },
        "description": {
          "short": `${args.options.description}`,
          "full": `${args.options.longDescription}`
        },
        "accentColor": args.options.accentColor || '#40497E',
        "isFullScreen": true,
        "staticTabs": [
          {
            "entityId": `sharepointportal_${appId}`,
            "name": `Portals-${args.options.name}`,
            "contentUrl": `https://${domain}/_layouts/15/teamslogon.aspx?spfx=true&dest=${appPortalUrl}`,
            "websiteUrl": portalUrl,
            "searchUrl": `https://${searchUrlPath}/_layouts/15/search.aspx?q={searchQuery}`,
            "scopes": ["personal"],
            "supportedPlatform": ["desktop"]
          }
        ],
        "permissions": [
          "identity",
          "messageTeamMembers"
        ],
        "validDomains": [
          domain,
          "*.login.microsoftonline.com",
          "*.sharepoint.com",
          "*.sharepoint-df.com",
          "spoppe-a.akamaihd.net",
          "spoprod-a.akamaihd.net",
          "resourceseng.blob.core.windows.net",
          "msft.spoppe.com"
        ],
        "webApplicationInfo": {
          "id": "00000003-0000-0ff1-ce00-000000000000",
          "resource": `https://${domain}`
        }
      };
      const manifestString = JSON.stringify(manifest, null, 2);

      try {
        // we need this to be able to inject mock AdmZip for testing
        /* c8 ignore next 3 */
        if (!this.archive) {
          this.archive = new AdmZip();
        }
        this.archive.addFile('manifest.json', Buffer.alloc(manifestString.length, manifestString, 'utf8'));
        this.archive.addLocalFile(coloredIconPath, undefined, coloredIconFileName);
        this.archive.addLocalFile(outlineIconPath, undefined, outlineIconFileName);
        this.archive.writeZip(`${args.options.name}.zip`);
      }
      catch (ex: any) {
        throw ex.message;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getWeb(args: CommandArgs, logger: Logger): Promise<WebProperties> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if site ${args.options.portalUrl} exists...`);
    }
    return await spo.getWeb(args.options.portalUrl, logger, this.verbose);
  }
}

export default new VivaConnectionsAppCreateCommand();