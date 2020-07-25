import * as fs from "fs";
import * as path from 'path';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandAction,
  CommandError
} from '../../../../Command';
import CdsProjectMutator from "../../cds-project-mutator";
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  path: string;
}

/*
 * Logic extracted from bolt.module.solution.dll
 * Version: 1.0.6
 * Class: bolt.module.solution.verbs.SolutionAddReferenceVerb
 */
class PaSolutionReferenceAddCommand extends Command {
  public get name(): string {
    return commands.SOLUTION_REFERENCE_ADD;
  }

  public get description(): string {
    return 'Adds a project reference to the solution in the current directory';
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    try {
      const referencedProjectFilePath: string = this.getSupportedProjectFiles(args.options.path)[0];
      const relativeReferencedProjectFilePath: string = path.relative(process.cwd(), referencedProjectFilePath);
      const cdsProjectFilePath: string = this.getCdsProjectFile(process.cwd())[0];
      const cdsProjectFileContent: string = fs.readFileSync(cdsProjectFilePath, 'utf8');

      const cdsProjectMutator = new CdsProjectMutator(cdsProjectFileContent);
      cdsProjectMutator.addProjectReference(relativeReferencedProjectFilePath);

      fs.writeFileSync(cdsProjectFilePath, cdsProjectMutator.cdsProjectDocument);

      cb();
    }
    catch (err) {
      cb(new CommandError(err));
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --path <path>',
        description: 'The path to the referenced project'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const existingCdsProjects: string[] = this.getCdsProjectFile(process.cwd());

      if (existingCdsProjects.length === 0) {
        return 'CDS solution project file with extension cdsproj was not found in the current directory.';
      }

      if (existingCdsProjects.length > 1) {
        return 'Multiple CDS solution project files with extension cdsproj were found in the current directory.';
      }

      if (!args.options.path) {
        return 'Missing required option path.';
      }

      if (!fs.existsSync(args.options.path)) {
        return `Path ${args.options.path} is not a valid path.`;
      }

      const existingSupportedProjects: string[] = this.getSupportedProjectFiles(args.options.path);
      if (existingSupportedProjects.length === 0) {
        return `No supported project type found in path ${args.options.path}.`;
      }

      if (existingSupportedProjects.length !== 1) {
        return `More than one supported project type found in path ${args.options.path}.`;
      }

      const cdsProjectName: string = path.parse(path.basename(existingCdsProjects[0])).name;
      const pcfProjectName: string = path.parse(path.basename(existingSupportedProjects[0])).name;

      if (cdsProjectName === pcfProjectName) {
        return `Not able to add reference to a project with same name as CDS project with name: ${pcfProjectName}.`;
      }

      return true;
    };
  }

  private getCdsProjectFile(rootPath: string): string[] {
    return fs.readdirSync(rootPath)
      .filter(fn => path.extname(fn).toLowerCase() === '.cdsproj')
      .map(entry => path.join(rootPath, entry));
  }

  private getSupportedProjectFiles(rootPath: string): string[] {
    return fs.readdirSync(rootPath).filter(fn => {
      const ext: string = path.extname(fn).toLowerCase();
      return ext === '.pcfproj' || ext === '.csproj';
    }).map(entry => path.join(rootPath, entry));
  }
}

module.exports = new PaSolutionReferenceAddCommand();