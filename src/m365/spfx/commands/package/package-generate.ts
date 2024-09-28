import AdmZip from 'adm-zip';
import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  allowTenantWideDeployment: boolean;
  developerMpnId?: string;
  developerName?: string;
  developerPrivacyUrl?: string;
  developerTermsOfUseUrl?: string;
  developerWebsiteUrl?: string;
  enableForTeams?: string;
  exposePageContextGlobally: boolean;
  exposeTeamsContextGlobally: boolean;
  html: string;
  name: string;
  webPartDescription: string;
  webPartTitle: string;
}

class SpfxPackageGenerateCommand extends AnonymousCommand {
  private static readonly enableForTeamsOptions: string[] = ['tab', 'personalApp', 'all'];
  private archive?: AdmZip;

  public get name(): string {
    return commands.PACKAGE_GENERATE;
  }

  public get description(): string {
    return 'Generates SharePoint Framework solution package with a no-framework web part rendering the specified HTML snippet';
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
        allowTenantWideDeployment: args.options.allowTenantWideDeployment === true,
        developerMpnId: typeof args.options.developerMpnId !== 'undefined',
        developerName: typeof args.options.developerName !== 'undefined',
        developerPrivacyUrl: typeof args.options.developerPrivacyUrl !== 'undefined',
        developerTermsOfUseUrl: typeof args.options.developerTermsOfUseUrl !== 'undefined',
        developerWebsiteUrl: typeof args.options.developerWebsiteUrl !== 'undefined',
        enableForTeams: args.options.enableForTeams,
        exposePageContextGlobally: args.options.exposePageContextGlobally === true,
        exposeTeamsContextGlobally: args.options.exposeTeamsContextGlobally === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-t, --webPartTitle <webPartTitle>' },
      { option: '-d, --webPartDescription <webPartDescription>' },
      { option: '-n, --name <name>' },
      { option: '--html <html>' },
      {
        option: '--enableForTeams [enableForTeams]',
        autocomplete: SpfxPackageGenerateCommand.enableForTeamsOptions
      },
      { option: '--exposePageContextGlobally' },
      { option: '--exposeTeamsContextGlobally' },
      { option: '--allowTenantWideDeployment' },
      { option: '--developerName [developerName]' },
      { option: '--developerPrivacyUrl [developerPrivacyUrl]' },
      { option: '--developerTermsOfUseUrl [developerTermsOfUseUrl]' },
      { option: '--developerWebsiteUrl [developerWebsiteUrl]' },
      { option: '--developerMpnId [developerMpnId]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.enableForTeams &&
          SpfxPackageGenerateCommand.enableForTeamsOptions.indexOf(args.options.enableForTeams) < 0) {
          return `${args.options.enableForTeams} is not a valid value for enableForTeams. Allowed values are: ${SpfxPackageGenerateCommand.enableForTeamsOptions.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const supportedHosts: string[] = ['SharePointWebPart'];
    if (args.options.enableForTeams === 'tab' || args.options.enableForTeams === 'all') {
      supportedHosts.push('TeamsTab');
    }
    if (args.options.enableForTeams === 'personalApp' || args.options.enableForTeams === 'all') {
      supportedHosts.push('TeamsPersonalApp');
    }

    const tokens: any = {
      clientSideAssetsFeatureId: SpfxPackageGenerateCommand.generateNewId(),
      developerName: args.options.developerName || 'Contoso',
      developerWebsiteUrl: args.options.developerWebsiteUrl || 'https://contoso.com/my-app',
      developerPrivacyUrl: args.options.developerPrivacyUrl || 'https://contoso.com/privacy',
      developerTermsOfUseUrl: args.options.developerTermsOfUseUrl || 'https://contoso.com/terms-of-use',
      developerMpnId: args.options.developerMpnId || '000000',
      exposePageContextGlobally: args.options.exposePageContextGlobally ? '!0' : '!1',
      exposeTeamsContextGlobally: args.options.exposeTeamsContextGlobally ? '!0' : '!1',
      html: args.options.html.replace(/"/g, '\\"').replace(/\r\n/g, ' ').replace(/\n/g, ' '),
      packageName: SpfxPackageGenerateCommand.getSafePackageName(args.options.webPartTitle),
      productId: SpfxPackageGenerateCommand.generateNewId(),
      skipFeatureDeployment: (args.options.allowTenantWideDeployment === true).toString(),
      supportedHosts: JSON.stringify(supportedHosts).replace(/"/g, '&quot;'),
      webPartId: SpfxPackageGenerateCommand.generateNewId(),
      webPartFeatureName: `${args.options.webPartTitle} Feature`,
      webPartFeatureDescription: `A feature which activates the Client-Side WebPart named ${args.options.webPartTitle}`,
      webPartAlias: SpfxPackageGenerateCommand.getWebPartAlias(args.options.webPartTitle),
      webPartName: args.options.webPartTitle,
      webPartSafeName: SpfxPackageGenerateCommand.getSafeWebPartName(args.options.webPartTitle),
      webPartDescription: args.options.webPartDescription,
      webPartModule: SpfxPackageGenerateCommand.getSafePackageName(args.options.webPartTitle)
    };

    let tmpDir: string | undefined = undefined;
    let error: any;
    try {
      if (this.verbose) {
        await logger.log(`Creating temp folder...`);
      }
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-spfx'));
      if (this.debug) {
        await logger.log(`Temp folder created at ${tmpDir}`);
      }

      if (this.verbose) {
        await logger.log('Copying files...');
      }
      const src: string = path.join(__dirname, 'package-generate', 'assets');
      fsUtil.copyRecursiveSync(src, tmpDir, s => SpfxPackageGenerateCommand.replaceTokens(s, tokens));

      const files: string[] = fsUtil.readdirR(tmpDir) as string[];
      if (this.verbose) {
        await logger.log('Processing files...');
      }
      for (const filePath of files) {
        if (this.debug) {
          await logger.log(`Processing ${filePath}...`);
        }

        if (!SpfxPackageGenerateCommand.isBinaryFile(filePath)) {
          if (this.verbose) {
            await logger.log('Replacing tokens...');
          }
          let fileContents: string = fs.readFileSync(filePath, 'utf-8');
          if (this.debug) {
            await logger.log('Before:');
            await logger.log(fileContents);
          }
          fileContents = SpfxPackageGenerateCommand.replaceTokens(fileContents, tokens);
          if (this.debug) {
            await logger.log('After:');
            await logger.log(fileContents);
          }
          fs.writeFileSync(filePath, fileContents, { encoding: 'utf-8' });
        }
        else {
          if (this.verbose) {
            await logger.log(`Binary file. Skipping replacing tokens in contents`);
          }
        }
      }

      if (this.verbose) {
        await logger.log('Creating .sppkg file...');
      }
      // we need this to be able to inject mock AdmZip for testing
      /* c8 ignore next 3 */
      if (!this.archive) {
        this.archive = new AdmZip();
      }
      const filesToZip: string[] = fsUtil.readdirR(tmpDir) as string[];
      for (const f of filesToZip) {
        if (this.debug) {
          await logger.log(`Adding ${f} to archive...`);
        }

        this.archive!.addLocalFile(f, path.relative(tmpDir as string, path.dirname(f)), path.basename(f));
      }
      if (this.debug) {
        await logger.log('Writing archive...');
      }
      this.archive.writeZip(`${args.options.name}.sppkg`);
    }
    catch (err: any) {
      error = err.message;
    }
    finally {
      try {
        if (tmpDir) {
          if (this.verbose) {
            await logger.log(`Deleting temp folder at ${tmpDir}...`);
          }
          fs.rmdirSync(tmpDir, { recursive: true });
        }
        if (error) {
          throw error;
        }
      }
      catch (ex: any) {
        if (ex === error) {
          throw ex;
        }

        throw `An error has occurred while removing the temp folder at ${tmpDir}. Please remove it manually.`;
      }
    }
  }

  private static replaceTokens(s: string, tokens: any): string {
    return s.replace(/\$([^\$]+)\$/g, (substring: string, token: string): string => {
      if (tokens[token]) {
        return tokens[token];
      }
      else {
        return substring;
      }
    });
  }

  private static isBinaryFile(filePath: string): boolean {
    return filePath.endsWith('.png');
  }

  private static getSafePackageName(packageName: string): string {
    return packageName.toLowerCase().replace(/[^a-zA-Z0-9]/g, '-');
  }

  private static getSafeWebPartName(webPartName: string): string {
    return webPartName.replace(/ /g, '-');
  }

  private static getWebPartAlias(webPartName: string): string {
    return 'AutoWP' + webPartName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 40);
  }

  private static generateNewId = (): string => {
    return v4();
  };
}

export default new SpfxPackageGenerateCommand();