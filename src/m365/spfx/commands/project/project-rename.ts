import * as fs from 'fs';
import * as path from 'path';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import { CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { BaseProjectCommand } from './base-project-command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  newName: string;
  generateNewId?: boolean;
}

class SpfxProjectRenameCommand extends BaseProjectCommand {
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;

  public get name(): string {
    return commands.PROJECT_RENAME;
  }

  public get description(): string {
    return 'Renames SharePoint Framework project';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.generateNewId = args.options.generateNewId;
    return telemetryProps;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --newName <newName>',
        description: 'New name for the project'
      },
      {
        option: '--generateNewId',
        description: 'Generate a new solution ID for the project'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectRenameCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    const packageJson: any = this.getProject(this.projectRootPath).packageJson;
    const projectName: string = packageJson.name;

    let newId: string = '';
    if (args.options.generateNewId) {
      newId = this.generateNewId();
      if (this.debug) {
        logger.logToStderr('Created new solution id');
        logger.logToStderr(newId);
      }
    }

    if (this.debug) {
      logger.logToStderr(`Renaming SharePoint Framework project to '${args.options.newName}'`);
    }

    try {
      this.replacePackageJsonContent(path.join(this.projectRootPath, 'package.json'), args, logger);
      this.replaceYoRcJsonContent(path.join(this.projectRootPath, '.yo-rc.json'), newId, args, logger);
      this.replacePackageSolutionJsonContent(path.join(this.projectRootPath, 'config', 'package-solution.json'), projectName, newId, args, logger);
      this.replaceDeployAzureStorageJsonContent(path.join(this.projectRootPath, 'config', 'deploy-azure-storage.json'), args, logger);
      this.replaceReadMeContent(path.join(this.projectRootPath, 'README.md'), projectName, args, logger);
    }
    catch (error) {
      cb(new CommandError(error));
      return;
    }

    if (this.verbose) {
      logger.logToStderr('DONE');
    }

    cb();
  }

  private generateNewId = (): string => {
    return v4();
  }

  private replacePackageJsonContent = (filePath: string, args: CommandArgs, logger: Logger) => {
    if (!fs.existsSync(filePath)) {
      return;
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const updatedContent = JSON.parse(existingContent);

    if (updatedContent &&
      updatedContent.name) {
      updatedContent.name = args.options.newName;
    }

    const updatedContentString: string = JSON.stringify(updatedContent, null, 2);

    if (updatedContentString !== existingContent) {
      fs.writeFileSync(filePath, updatedContentString, 'utf-8');

      if (this.debug) {
        logger.logToStderr(`Updated ${path.basename(filePath)}`);
      }
    }
  }

  private replaceYoRcJsonContent = (filePath: string, newId: string, args: CommandArgs, logger: Logger) => {
    if (!fs.existsSync(filePath)) {
      return;
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const updatedContent = JSON.parse(existingContent);

    if (updatedContent &&
      updatedContent['@microsoft/generator-sharepoint'] &&
      updatedContent['@microsoft/generator-sharepoint'].libraryName) {
      updatedContent['@microsoft/generator-sharepoint'].libraryName = args.options.newName;
    }
    if (updatedContent &&
      updatedContent['@microsoft/generator-sharepoint'] &&
      updatedContent['@microsoft/generator-sharepoint'].solutionName) {
      updatedContent['@microsoft/generator-sharepoint'].solutionName = args.options.newName;
    }
    if (updatedContent &&
      updatedContent['@microsoft/generator-sharepoint'] &&
      updatedContent['@microsoft/generator-sharepoint'].libraryId &&
      args.options.generateNewId) {
      updatedContent['@microsoft/generator-sharepoint'].libraryId = newId;
    }

    const updatedContentString: string = JSON.stringify(updatedContent, null, 2);

    if (updatedContentString !== existingContent) {
      fs.writeFileSync(filePath, updatedContentString, 'utf-8');

      if (this.debug) {
        logger.logToStderr(`Updated ${path.basename(filePath)}`);
      }
    }
  }

  private replacePackageSolutionJsonContent = (filePath: string, projectName: string, newId: string, args: CommandArgs, logger: Logger) => {
    if (!fs.existsSync(filePath)) {
      return;
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const updatedContent = JSON.parse(existingContent);

    if (updatedContent &&
      updatedContent.solution &&
      updatedContent.solution.name) {
      updatedContent.solution.name = updatedContent.solution.name.replace(new RegExp(projectName, 'g'), args.options.newName);
    }
    if (updatedContent &&
      updatedContent.solution &&
      updatedContent.solution.id &&
      args.options.generateNewId) {
      updatedContent.solution.id = newId;
    }
    if (updatedContent &&
      updatedContent.paths &&
      updatedContent.paths.zippedPackage) {
      updatedContent.paths.zippedPackage = updatedContent.paths.zippedPackage.replace(new RegExp(projectName, 'g'), args.options.newName);
    }

    const updatedContentString: string = JSON.stringify(updatedContent, null, 2);

    if (updatedContentString !== existingContent) {
      fs.writeFileSync(filePath, updatedContentString, 'utf-8');

      if (this.debug) {
        logger.logToStderr(`Updated ${path.basename(filePath)}`);
      }
    }
  }

  private replaceDeployAzureStorageJsonContent = (filePath: string, args: CommandArgs, logger: Logger) => {
    if (!fs.existsSync(filePath)) {
      return;
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const updatedContent = JSON.parse(existingContent);

    if (updatedContent &&
      updatedContent.container) {
      updatedContent.container = args.options.newName;
    }

    const updatedContentString: string = JSON.stringify(updatedContent, null, 2);

    if (updatedContentString !== existingContent) {
      fs.writeFileSync(filePath, updatedContentString, 'utf-8');

      if (this.debug) {
        logger.logToStderr(`Updated ${path.basename(filePath)}`);
      }
    }
  }

  private replaceReadMeContent = (filePath: string, projectName: string, args: CommandArgs, logger: Logger) => {
    if (!fs.existsSync(filePath)) {
      return;
    }

    const existingContent: string = fs.readFileSync(filePath, 'utf-8');
    const updatedContent = existingContent.replace(new RegExp(projectName, 'g'), args.options.newName);

    if (updatedContent !== existingContent) {
      fs.writeFileSync(filePath, updatedContent, 'utf-8');

      if (this.debug) {
        logger.logToStderr(`Updated ${path.basename(filePath)}`);
      }
    }
  }
}

module.exports = new SpfxProjectRenameCommand();