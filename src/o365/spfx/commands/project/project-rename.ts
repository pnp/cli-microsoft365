import commands from '../../commands';
import Command, {
  CommandOption, CommandAction, CommandError, CommandValidate
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { BaseProjectCommand } from './base-project-command';
import * as path from 'path';
import * as fs from 'fs';
const uuid = require('uuid');

const vorpal: Vorpal = require('../../../../vorpal-init');

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
        description: 'Generate a new solution id for the project'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.newName) {
        return 'Required parameter newName missing';
      }

      return true;
    };
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      args = (cmd as any).processArgs(args);
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectRenameCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    const packageJson: any = this.getProject(this.projectRootPath).packageJson;
    const projectName: string = packageJson.name;

    let newId: string;
    if (args.options.generateNewId) {
      newId = uuid.v4();
      if (this.debug) {
        cmd.log('Created new solution id');
        cmd.log(newId);
      }
    }

    const replacePackageJsonContent = (filePath: string) => {
      if (!fs.existsSync(filePath)) {
        return;
      }
      const existingContent = fs.readFileSync(filePath, 'utf-8');
      const updatedContent = JSON.parse(existingContent);
      if (updatedContent &&
        updatedContent.name) {
        updatedContent.name = args.options.newName;
      }
      fs.writeFileSync(filePath, JSON.stringify(updatedContent, null, 2), 'utf-8');
      if (this.debug) {
        cmd.log(`Updated ${filePath.split('/').pop()}`);
      }
    }

    const replaceYoRcJsonContent = (filePath: string) => {
      if (!fs.existsSync(filePath)) {
        return;
      }
      const existingContent = fs.readFileSync(filePath, 'utf-8');
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
      fs.writeFileSync(filePath, JSON.stringify(updatedContent, null, 2), 'utf-8');
      if (this.debug) {
        cmd.log(`Updated ${filePath.split('/').pop()}`);
      }
    }

    const replacePackageSolutionJsonContent = (filePath: string) => {
      if (!fs.existsSync(filePath)) {
        return;
      }
      const existingContent = fs.readFileSync(filePath, 'utf-8');
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
      fs.writeFileSync(filePath, JSON.stringify(updatedContent, null, 2), 'utf-8');
      if (this.debug) {
        cmd.log(`Updated ${filePath.split('/').pop()}`);
      }
    }

    const replaceDeployAzureStorageJsonContent = (filePath: string) => {
      if (!fs.existsSync(filePath)) {
        return;
      }
      const existingContent = fs.readFileSync(filePath, 'utf-8');
      const updatedContent = JSON.parse(existingContent);
      if (updatedContent &&
        updatedContent.container) {
        updatedContent.container = args.options.newName;
      }
      fs.writeFileSync(filePath, JSON.stringify(updatedContent, null, 2), 'utf-8');
      if (this.debug) {
        cmd.log(`Updated ${filePath.split('/').pop()}`);
      }
    }

    const replaceReadMeContent = (filePath: string) => {
      if (!fs.existsSync(filePath)) {
        return;
      }
      const existingContent = fs.readFileSync(filePath, 'utf-8');
      const updatedContent = existingContent.replace(new RegExp(projectName, 'g'), args.options.newName);
      fs.writeFileSync(filePath, updatedContent, 'utf-8');
      if (this.debug) {
        cmd.log(`Updated ${filePath.split('/').pop()}`);
      }
    }

    if (this.debug) {
      cmd.log(`Renaming SharePoint Framework project to '${args.options.newName}'`);
    }

    replacePackageJsonContent(path.join(this.projectRootPath, 'package.json'));
    replaceYoRcJsonContent(path.join(this.projectRootPath, '.yo-rc.json'));
    replacePackageSolutionJsonContent(path.join(this.projectRootPath, 'config/package-solution.json'));
    replaceDeployAzureStorageJsonContent(path.join(this.projectRootPath, 'config/deploy-azure-storage.json'));
    replaceReadMeContent(path.join(this.projectRootPath, 'README.md'));

    if (this.verbose) {
      cmd.log('DONE');
    }

    cb();
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROJECT_RENAME).helpInformation());
    log(
      `   ${chalk.yellow('Important:')} Run this command in the folder where the project for which you
      want to externalize dependencies is located.

  Examples:

    Renames SharePoint Framework project to contoso
      ${commands.PROJECT_RENAME} --newName contoso

    Renames SharePoint Framework project to contoso with new solution Id
      ${commands.PROJECT_RENAME} --newName contoso --generateNewId
`);
  }
}

module.exports = new SpfxProjectRenameCommand();