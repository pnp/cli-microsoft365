import * as chalk from 'chalk';
import * as fs from "fs";
import * as path from 'path';
import { v4 } from 'uuid';
import { Logger } from "../../../../cli";
import {
  CommandError, CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from '../../commands';
import TemplateInstantiator from "../../template-instantiator";
import { SolutionInitVariables } from "./solution-init/solution-init-variables";

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
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
        logger.logToStderr(`publisherName: ${publisherName}`);
        logger.logToStderr(`publisherPrefix: ${publisherPrefix}`);
        logger.logToStderr(`customizationOptionValuePrefix: ${customizationOptionValuePrefix}`);
        logger.logToStderr(`solutionInitTemplatePath: ${solutionInitTemplatePath}`);
        logger.logToStderr(`cdsAssetsTemplatePath: ${cdsAssetsTemplatePath}`);
        logger.logToStderr(`workingDirectory: ${workingDirectory}`);
        logger.logToStderr(`workingDirectoryName: ${workingDirectoryName}`);
        logger.logToStderr(`cdsAssetsDirectory: ${cdsAssetsDirectory}`);
        logger.logToStderr(`cdsAssetsDirectorySolutionsFile: ${cdsAssetsDirectorySolutionsFile}`);
      }

      TemplateInstantiator.instantiate(logger, solutionInitTemplatePath, workingDirectory, false, variables, this.verbose);

      if (this.verbose) {
        logger.logToStderr(` `);
      }

      logger.log(chalk.green(`CDS solution project with name '${workingDirectoryName}' created successfully in current directory.`));

      const cdsAssetsExist: boolean = fs.existsSync(cdsAssetsDirectory) && fs.existsSync(cdsAssetsDirectorySolutionsFile);
      if (cdsAssetsExist) {
        logger.log(chalk.yellow(`CDS solution files already exist in the current directory. Skipping CDS solution files creation.`));
      }
      else {
        TemplateInstantiator.instantiate(logger, cdsAssetsTemplatePath, cdsAssetsDirectory, false, variables, this.verbose);
        logger.log(chalk.green(`CDS solution files were successfully created for this project in the sub-directory 'Other', using solution name '${workingDirectory}', publisher name '${publisherName}', and customization prefix '${publisherPrefix}'.`));
        logger.log(`Please verify the publisher information and solution name found in the '${chalk.grey('Solution.xml')}' file.`);
      }

      cb();
    }
    catch (err: any) {
      cb(new CommandError(err));
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--publisherName <publisherName>'
      },
      {
        option: '--publisherPrefix <publisherPrefix>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (fs.readdirSync(process.cwd()).some(fn => path.extname(fn).toLowerCase() === '.cdsproj')) {
      return 'CDS project creation failed. The current directory already contains a project. Please create a new directory and retry the operation.';
    }

    const workingDirectoryName: string = path.basename(process.cwd());
    if (!Utils.isValidFileName(workingDirectoryName)) {
      return `Empty or invalid project name '${workingDirectoryName}'`;
    }

    if (args.options.publisherPrefix) {
      if (args.options.publisherPrefix.length < 2 || args.options.publisherPrefix.length > 8 || !/^(?!mscrm)^([a-zA-Z])\w*$/i.test(args.options.publisherPrefix)) {
        return `Value of 'publisherPrefix' is invalid. The prefix must be 2 to 8 characters long, can only consist of alpha-numerics, must start with a letter, and cannot start with 'mscrm'.`;
      }
    }
    else {
      return 'Missing required option publisherPrefix.';
    }

    if (args.options.publisherName) {
      if (!/^([a-zA-Z_])\w*$/i.test(args.options.publisherName)) {
        return `Value of 'publisherName' is invalid. Only characters within the ranges [A-Z], [a-z], [0-9], or _ are allowed. The first character may only be in the ranges [A-Z], [a-z], or _.`;
      }
    }
    else {
      return 'Missing required option publisherName.';
    }

    return true;
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

module.exports = new PaSolutionInitCommand();