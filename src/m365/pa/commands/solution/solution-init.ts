import chalk from 'chalk';
import fs from 'fs';
import path from 'path';
import url from 'url';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
import TemplateInstantiator from '../../template-instantiator.js';
import { SolutionInitVariables } from './solution-init/solution-init-variables.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  publisherName: string;
  publisherPrefix: string;
}

/*
 * Logic extracted from bolt.module.solution.dll
 * Version: 1.0.6
 * Class: bolt.module.solution.SolutionInitVerb
 */
class PaSolutionInitCommand extends AnonymousCommand {
  public get name(): string {
    return commands.SOLUTION_INIT;
  }

  public get description(): string {
    return 'Initializes a directory with a new CDS solution project';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--publisherName <publisherName>'
      },
      {
        option: '--publisherPrefix <publisherPrefix>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (fs.readdirSync(process.cwd()).some(fn => path.extname(fn).toLowerCase() === '.cdsproj')) {
          return 'CDS project creation failed. The current directory already contains a project. Please create a new directory and retry the operation.';
        }

        const workingDirectoryName: string = path.basename(process.cwd());
        if (!validation.isValidFileName(workingDirectoryName)) {
          return `Empty or invalid project name '${workingDirectoryName}'`;
        }

        if (args.options.publisherPrefix.length < 2 || args.options.publisherPrefix.length > 8 || !/^(?!mscrm)^([a-zA-Z])\w*$/i.test(args.options.publisherPrefix)) {
          return `Value of 'publisherPrefix' is invalid. The prefix must be 2 to 8 characters long, can only consist of alpha-numerics, must start with a letter, and cannot start with 'mscrm'.`;
        }

        if (!/^([a-zA-Z_])\w*$/i.test(args.options.publisherName)) {
          return `Value of 'publisherName' is invalid. Only characters within the ranges [A-Z], [a-z], [0-9], or _ are allowed. The first character may only be in the ranges [A-Z], [a-z], or _.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const solutionInitTemplatePath: string = path.join(__dirname, 'solution-init', 'assets');
      const cdsAssetsTemplatePath: string = path.join(solutionInitTemplatePath, 'Other');
      const workingDirectory: string = process.cwd();
      const workingDirectoryName: string = path.basename(workingDirectory);
      const cdsAssetsDirectory: string = path.join(workingDirectory, 'Other');
      const cdsAssetsDirectorySolutionsFile: string = path.join(workingDirectory, 'Other', 'Solution.xml');
      const publisherName: string = args.options.publisherName;
      const publisherPrefix: string = args.options.publisherPrefix.toLocaleLowerCase();
      const customizationOptionValuePrefix: string = this.generateOptionValuePrefixForPublisher(publisherPrefix);
      const variables: SolutionInitVariables = {
        "$publisherName$": publisherName,
        "$customizationPrefix$": publisherPrefix,
        "$customizationOptionValuePrefix$": customizationOptionValuePrefix,
        "$cdsProjectGuid$": v4(),
        "solutionprojecttype": workingDirectoryName,
        "$solutionName$": workingDirectoryName
      };

      if (this.verbose) {
        await logger.logToStderr(`publisherName: ${publisherName}`);
        await logger.logToStderr(`publisherPrefix: ${publisherPrefix}`);
        await logger.logToStderr(`customizationOptionValuePrefix: ${customizationOptionValuePrefix}`);
        await logger.logToStderr(`solutionInitTemplatePath: ${solutionInitTemplatePath}`);
        await logger.logToStderr(`cdsAssetsTemplatePath: ${cdsAssetsTemplatePath}`);
        await logger.logToStderr(`workingDirectory: ${workingDirectory}`);
        await logger.logToStderr(`workingDirectoryName: ${workingDirectoryName}`);
        await logger.logToStderr(`cdsAssetsDirectory: ${cdsAssetsDirectory}`);
        await logger.logToStderr(`cdsAssetsDirectorySolutionsFile: ${cdsAssetsDirectorySolutionsFile}`);
      }

      TemplateInstantiator.instantiate(logger, solutionInitTemplatePath, workingDirectory, false, variables, this.verbose);

      if (this.verbose) {
        await logger.logToStderr(` `);
      }

      await logger.log(chalk.green(`CDS solution project with name '${workingDirectoryName}' created successfully in current directory.`));

      const cdsAssetsExist: boolean = fs.existsSync(cdsAssetsDirectory) && fs.existsSync(cdsAssetsDirectorySolutionsFile);
      if (cdsAssetsExist) {
        await logger.log(chalk.yellow(`CDS solution files already exist in the current directory. Skipping CDS solution files creation.`));
      }
      else {
        TemplateInstantiator.instantiate(logger, cdsAssetsTemplatePath, cdsAssetsDirectory, false, variables, this.verbose);
        await logger.log(chalk.green(`CDS solution files were successfully created for this project in the sub-directory 'Other', using solution name '${workingDirectory}', publisher name '${publisherName}', and customization prefix '${publisherPrefix}'.`));
        await logger.log(`Please verify the publisher information and solution name found in the '${chalk.grey('Solution.xml')}' file.`);
      }
    }
    catch (err: any) {
      throw err;
    }
  }

  private generateOptionValuePrefixForPublisher(customizationPrefix: string): string {
    if (customizationPrefix.toLocaleLowerCase() !== "new") {
      return this.generateOptionValuePrefixForPublisherInternal(this.getHashCode(customizationPrefix));
    }

    return "10000";
  }

  private generateOptionValuePrefixForPublisherInternal(customizationPrefixHashCode: number): string {
    return Math.abs(customizationPrefixHashCode % 90000) + 10000 + "";
  }

  private getHashCode(s: string): number {
    let h = 0;
    for (let i = 0; i < s.length; i++) {
      h = Math.imul(31, h) + s.charCodeAt(i) | 0;
    }

    return h;
  }
}

export default new PaSolutionInitCommand();