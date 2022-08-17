import * as fs from "fs";
import * as path from 'path';
import { Logger } from "../../../../cli";
import {
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from "../../../base/AnonymousCommand";
import CdsProjectMutator from "../../cds-project-mutator";
import commands from '../../commands';

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
class PaSolutionReferenceAddCommand extends AnonymousCommand {
  public get name(): string {
    return commands.SOLUTION_REFERENCE_ADD;
  }

  public get description(): string {
    return 'Adds a project reference to the solution in the current directory';
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --path <path>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const existingCdsProjects: string[] = this.getCdsProjectFile(process.cwd());

        if (existingCdsProjects.length === 0) {
          return 'CDS solution project file with extension cdsproj was not found in the current directory.';
        }
    
        if (existingCdsProjects.length > 1) {
          return 'Multiple CDS solution project files with extension cdsproj were found in the current directory.';
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
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    try {
      const referencedProjectFilePath: string = this.getSupportedProjectFiles(args.options.path)[0];
      const relativeReferencedProjectFilePath: string = path.relative(process.cwd(), referencedProjectFilePath);
      const cdsProjectFilePath: string = this.getCdsProjectFile(process.cwd())[0];
      const cdsProjectFileContent: string = fs.readFileSync(cdsProjectFilePath, 'utf8');

      const cdsProjectMutator = new CdsProjectMutator(cdsProjectFileContent);
      cdsProjectMutator.addProjectReference(relativeReferencedProjectFilePath);

      fs.writeFileSync(cdsProjectFilePath, cdsProjectMutator.cdsProjectDocument as any);

      cb();
    }
    catch (err: any) {
      cb(new CommandError(err));
    }
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