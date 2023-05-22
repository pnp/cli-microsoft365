import * as AdmZip from 'adm-zip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import chalk = require('chalk');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  fix?: boolean;
  supportedHost?: string;
}

class SpfxPackageTeamsEnable extends AnonymousCommand {
  private static readonly allowedSupportedHosts: string[] = ['TeamsPersonalApp', 'TeamsMeetingApp', 'TeamsTab'];
  private solutionZip?: AdmZip;
  private fixZip?: AdmZip;

  public get name(): string {
    return commands.PACKAGE_TEAMS_ENABLE;
  }

  public get description(): string {
    return 'Checks if the SPFx package is enabled for deployment in Teams. Optionally, changes the package to enable it for deployment in Teams';
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
        fix: args.options.fix,
        supportedHost: typeof args.options.supportedHost !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --filePath <filePath>'
      },
      {
        option: '--fix'
      },
      {
        option: '--supportedHost [--supportedHost]',
        autocomplete: SpfxPackageTeamsEnable.allowedSupportedHosts
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const fullPath: string = path.resolve(args.options.filePath);

        if (!fs.existsSync(fullPath)) {
          return `Specified sppkg file with path '${fullPath}' does not exist`;
        }

        if (!fullPath.endsWith('.sppkg')) {
          return `Specified file is not of the valid file type. Please specify a valid sppkg file`;
        }

        if (args.options.supportedHost && args.options.supportedHost.split(',').some(splittedHost => !SpfxPackageTeamsEnable.allowedSupportedHosts.includes(splittedHost))) {
          return `The supportedHost contains an invalid value. Possible supportedHost values are: ${SpfxPackageTeamsEnable.allowedSupportedHosts.join(',')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let tmpDir: string | undefined = undefined;

    try {
      if (this.verbose) {
        logger.logToStderr('Creating temp folder to save the sppkg content in');
      }

      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-spfx'));

      if (this.verbose) {
        logger.logToStderr(`Temp folder created at path: ${tmpDir}`);
      }

      const fullPath: string = path.resolve(args.options.filePath);
      // we need this to be able to inject mock AdmZip for testing
      /* c8 ignore next 3 */
      if (!this.solutionZip) {
        this.solutionZip = new AdmZip(fullPath);
      }

      this.solutionZip.extractAllTo(tmpDir);
      const files = fs.readdirSync(tmpDir);

      let fixesApplied = false;

      files.forEach(file => {

        if (validation.isValidGuid(file)) {

          if (this.verbose) {
            logger.logToStderr(`Found webpart folder with id ${file}`);
          }

          const dirLocation = `${tmpDir}\\${file}`;
          const webpartFiles = fs.readdirSync(dirLocation);

          webpartFiles.forEach((wpFile: string) => {
            if (this.verbose) {
              logger.logToStderr(`Reading contents for file ${wpFile}`);
            }

            const fileLocation = `${dirLocation}\\${wpFile}`;
            let fileContent = fs.readFileSync(fileLocation, { encoding: 'utf-8' });
            const regex = new RegExp('ComponentManifest=\\"[^\\"]+\\"');

            if (regex.test(fileContent)) {

              const matches = fileContent.match(regex);
              const componentManifest = matches![0];
              const componentManifestReplaced = componentManifest.replace(/&quot;/gi, '"').replace('ComponentManifest=\"', '').slice(0, -1);
              const parsedComponentManifest: { id: string, alias: string, supportedHosts: string[] } = JSON.parse(componentManifestReplaced);
              const supportedHostMatches = parsedComponentManifest.supportedHosts.filter(supHost => supHost.startsWith('Teams'));

              if (supportedHostMatches.length === 0) {
                logger.logToStderr(chalk.red(`Webpart with id ${parsedComponentManifest.id} and alias ${parsedComponentManifest.alias} is not set-up as a Teams app.`));

                if (args.options.fix) {

                  if (this.verbose) {
                    logger.logToStderr('Time to fix the webpart to make it possible to set up as a Teams app.');
                  }

                  if (args.options.supportedHost) {
                    args.options.supportedHost.split(',').forEach((supportedHost: string) => {
                      if (!parsedComponentManifest.supportedHosts.some(existing => existing === supportedHost)) {
                        parsedComponentManifest.supportedHosts.push(supportedHost);
                      }
                    });
                  }
                  else {
                    parsedComponentManifest.supportedHosts.push('TeamsPersonalApp');
                  }

                  const revertReplace = JSON.stringify(parsedComponentManifest).replace(/["]+/g, '&quot;');
                  fileContent = fileContent.replace(componentManifest, `ComponentManifest="${revertReplace}"`);
                  fs.writeFileSync(fileLocation, fileContent);
                  fixesApplied = true;
                }
              }
              else {
                logger.logToStderr(chalk.green(`Webpart with id ${parsedComponentManifest.id} and alias ${parsedComponentManifest.alias} is set-up as a Teams app. Supported hosts: ${supportedHostMatches.join(', ')}`));
              }
            }
          });
        }
      });

      if (fixesApplied) {
        // we need this to be able to inject mock AdmZip for testing
        /* c8 ignore next 3 */
        if (!this.fixZip) {
          this.fixZip = new AdmZip();
        }
        this.fixZip.addLocalFolder(tmpDir);
        fs.writeFileSync(fullPath, this.fixZip.toBuffer());
      }

    }
    catch (err: any) {
      this.handleError(err);
    }
    finally {
      if (tmpDir) {
        fs.rmdirSync(tmpDir, { recursive: true });
      }
    }
  }
}

module.exports = new SpfxPackageTeamsEnable();