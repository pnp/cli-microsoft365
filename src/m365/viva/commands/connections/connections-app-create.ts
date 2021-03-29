import * as AdmZip from 'adm-zip';
import * as fs from 'fs';
import * as path from 'path';
import { v4 } from 'uuid';
import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandErrorWithOutput, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import * as spoWebGetCommand from '../../../spo/commands/web/web-get';
import { Options as SpoWebGetCommandOptions } from '../../../spo/commands/web/web-get';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  accentColor?: string;
  appName: string;
  coloredIconPath: string;
  companyName: string;
  companyWebsiteUrl: string;
  description: string;
  force?: boolean;
  longDescription: string;
  outlineIconPath: string;
  portalUrl: string;
  privacyPolicyUrl?: string;
  termsOfUseUrl?: string;
}

class VivaConnectionsAppCreateCommand extends AnonymousCommand {
  private archive?: AdmZip;

  public get name(): string {
    return commands.VIVA_CONNECTIONS_APP_CREATE;
  }

  public get description(): string {
    return 'Creates Viva Connections app';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getWeb(args, logger)
      .then((getWebOutput: CommandOutput): void => {
        if (this.debug) {
          logger.logToStderr(getWebOutput.stderr);
        }

        if (this.verbose) {
          logger.logToStderr(`Site found at ${args.options.portalUrl}. Checking if it's a communication site...`);
        }

        const web: {
          Configuration: number;
          WebTemplate: string;
        } = JSON.parse(getWebOutput.stdout);

        if (web.WebTemplate !== 'SITEPAGEPUBLISHING' ||
          web.Configuration !== 0) {
          return cb(`Site ${args.options.portalUrl} is not a Communication Site. Please specify a different site and try again.`);
        }

        if (this.verbose) {
          logger.logToStderr(`Site ${args.options.portalUrl} is a Communication Site. Building app...`);
        }

        const portalUrl: URL = new URL(args.options.portalUrl);
        const appPortalUrl: string = `${args.options.portalUrl}${args.options.portalUrl.indexOf('?') > -1 ? '&' : '?'}app=portals`;
        let searchUrlPath: string = portalUrl.hostname;
        if (portalUrl.pathname.indexOf('/teams') > -1 || portalUrl.pathname.indexOf('/sites') > -1) {
          const firstTwoUrlSegments = portalUrl.pathname.match(/^\/[^\/]+\/[^\/]+/);
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
          "packageName": `com.microsoft.teams.${args.options.appName}`,
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
            "short": args.options.appName,
            "full": args.options.appName
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
              "name": `Portals-${args.options.appName}`,
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
          this.archive.writeZip(`${args.options.appName}.zip`);
          cb();
        }
        catch (ex) {
          cb(ex.message);
        }
      }, (err: CommandErrorWithOutput) => {
        if (this.debug) {
          logger.logToStderr(err.stderr);
        }

        cb(err.error);
      });
  }

  private getWeb(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Checking if site ${args.options.url} exists...`);
    }

    const options: SpoWebGetCommandOptions = {
      webUrl: args.options.portalUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };
    return Cli.executeCommandWithOutput(spoWebGetCommand as Command, { options: { ...options, _: [] } });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--portalUrl <portalUrl>' },
      { option: '--appName <appName>' },
      { option: '--description <description>' },
      { option: '--longDescription <longDescription>' },
      { option: '--privacyPolicyUrl [privacyPolicyUrl]' },
      { option: '--termsOfUseUrl [termsOfUseUrl]' },
      { option: '--companyName <companyName>' },
      { option: '--companyWebsiteUrl <companyWebsiteUrl>' },
      { option: '--coloredIconPath <coloredIconPath>' },
      { option: '--outlineIconPath <outlineIconPath>' },
      { option: '--accentColor [accentColor]' },
      { option: '--force' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.appName.length > 30) {
      return `App name must not exceed 30 characters`;
    }

    if (args.options.description &&
      args.options.description.length > 80) {
      return 'Description must not exceed 80 characters';
    }

    if (args.options.longDescription &&
      args.options.longDescription.length > 4000) {
      return 'Long description must not exceed 4000 characters';
    }

    const appFilePath = path.resolve(`${args.options.appName}.zip`);
    if (fs.existsSync(appFilePath) && !args.options.force) {
      return `File ${appFilePath} already exists. Delete the file or use the --force option to overwrite the existing file`;
    }

    const coloredIconPath = path.resolve(args.options.coloredIconPath);
    if (!fs.existsSync(coloredIconPath)) {
      return `File ${coloredIconPath} doesn't exist`;
    }

    const outlineIconPath = path.resolve(args.options.outlineIconPath);
    if (!fs.existsSync(outlineIconPath)) {
      return `File ${outlineIconPath} doesn't exist`;
    }

    return true;
  }
}

module.exports = new VivaConnectionsAppCreateCommand();