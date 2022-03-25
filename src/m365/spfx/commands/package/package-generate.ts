import * as AdmZip from 'adm-zip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { fsUtil } from '../../../../utils';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

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
  packageName: string;
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.allowTenantWideDeployment = args.options.allowTenantWideDeployment === true;
    telemetryProps.developerMpnId = typeof args.options.developerMpnId !== 'undefined';
    telemetryProps.developerName = typeof args.options.developerName !== 'undefined';
    telemetryProps.developerPrivacyUrl = typeof args.options.developerPrivacyUrl !== 'undefined';
    telemetryProps.developerTermsOfUseUrl = typeof args.options.developerTermsOfUseUrl !== 'undefined';
    telemetryProps.developerWebsiteUrl = typeof args.options.developerWebsiteUrl !== 'undefined';
    telemetryProps.enableForTeams = args.options.enableForTeams;
    telemetryProps.exposePageContextGlobally = args.options.exposePageContextGlobally === true;
    telemetryProps.exposeTeamsContextGlobally = args.options.exposeTeamsContextGlobally === true;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
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
        logger.log(`Creating temp folder...`);
      }
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-spfx'));
      if (this.debug) {
        logger.log(`Temp folder created at ${tmpDir}`);
      }

      if (this.verbose) {
        logger.log('Copying files...');
      }
      const src: string = path.join(__dirname, 'package-generate', 'assets');
      fsUtil.copyRecursiveSync(src, tmpDir, s => SpfxPackageGenerateCommand.replaceTokens(s, tokens));

      const files: string[] = fsUtil.readdirR(tmpDir) as string[];
      if (this.verbose) {
        logger.log('Processing files...');
      }
      files.forEach(filePath => {
        if (this.debug) {
          logger.log(`Processing ${filePath}...`);
        }

        if (!SpfxPackageGenerateCommand.isBinaryFile(filePath)) {
          if (this.verbose) {
            logger.log('Replacing tokens...');
          }
          let fileContents: string = fs.readFileSync(filePath, 'utf-8');
          if (this.debug) {
            logger.log('Before:');
            logger.log(fileContents);
          }
          fileContents = SpfxPackageGenerateCommand.replaceTokens(fileContents, tokens);
          if (this.debug) {
            logger.log('After:');
            logger.log(fileContents);
          }
          fs.writeFileSync(filePath, fileContents, { encoding: 'utf-8' });
        }
        else {
          if (this.verbose) {
            logger.log(`Binary file. Skipping replacing tokens in contents`);
          }
        }
      });

      if (this.verbose) {
        logger.log('Creating .sppkg file...');
      }
      // we need this to be able to inject mock AdmZip for testing
      /* c8 ignore next 3 */
      if (!this.archive) {
        this.archive = new AdmZip();
      }
      const filesToZip: string[] = fsUtil.readdirR(tmpDir) as string[];
      filesToZip.forEach(f => {
        if (this.debug) {
          logger.log(`Adding ${f} to archive...`);
        }
        this.archive!.addLocalFile(f, path.relative(tmpDir as string, path.dirname(f)), path.basename(f));
      });
      if (this.debug) {
        logger.log('Writing archive...');
      }
      this.archive.writeZip(`${args.options.packageName}.sppkg`);
    }
    catch (err: any) {
      error = err.message;
    }
    finally {
      try {
        if (tmpDir) {
          if (this.verbose) {
            logger.log(`Deleting temp folder at ${tmpDir}...`);
          }
          fs.rmdirSync(tmpDir, { recursive: true });
        }
        cb(error);
      }
      catch (ex) {
        cb(`An error has occurred while removing the temp folder at ${tmpDir}. Please remove it manually.`);
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
    return 'AutoWP' + webPartName.replace(/[^a-zA-Z0-9]/g, '').substr(0, 40);
  }

  private static generateNewId = (): string => {
    return v4();
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-t, --webPartTitle <webPartTitle>' },
      { option: '-d, --webPartDescription <webPartDescription>' },
      { option: '-n, --packageName <packageName>' },
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.enableForTeams &&
      SpfxPackageGenerateCommand.enableForTeamsOptions.indexOf(args.options.enableForTeams) < 0) {
      return `${args.options.enableForTeams} is not a valid value for enableForTeams. Allowed values are: ${SpfxPackageGenerateCommand.enableForTeamsOptions.join(', ')}`;
    }

    return true;
  }
}

module.exports = new SpfxPackageGenerateCommand();