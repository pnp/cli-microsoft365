import * as AdmZip from 'adm-zip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Logger } from '../../../../../cli/Logger';
import GlobalOptions from '../../../../../GlobalOptions';
import AnonymousCommand from '../../../../base/AnonymousCommand';
import commands from '../../../commands';
import { validation } from '../../../../../utils/validation';
import chalk = require('chalk');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  fix?: boolean;
}

class SpfxPackageTeamsEnable extends AnonymousCommand {
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
        fix: args.options.fix
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-f, --filePath <filePath>' },
      { option: '--fix' },
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

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let tmpDir: string | undefined = undefined;
    try {
      if (this.verbose) {
        logger.logToStderr('Creating temp folder...');
      }
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-spfx'));
      if (this.verbose) {
        logger.logToStderr(`Temp folder created at path: ${tmpDir}`);
      }
      const solutionZip = new AdmZip(args.options.filePath);
      solutionZip.extractAllTo(tmpDir);
      const files = fs.readdirSync(tmpDir);
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
            const fileContent = fs.readFileSync(fileLocation, { encoding: 'utf-8' });
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
                    logger.logToStderr('Time to fix the webpart to make it possible to set up as a Teams app');
                  }
                }
              }
              else {
                logger.logToStderr(chalk.green(`Webpart with id ${parsedComponentManifest.id} and alias ${parsedComponentManifest.alias} is set-up as a Teams app. Supported hosts: ${supportedHostMatches.join(', ')}`));
              }
            }
          });
        }
      });
    }
    catch {

    }
    finally {
      if (tmpDir) {
        logger.log(tmpDir);
        //fs.rmdirSync(tmpDir, { recursive: true });
      }
    }
    return;
  }
}

module.exports = new SpfxPackageTeamsEnable();