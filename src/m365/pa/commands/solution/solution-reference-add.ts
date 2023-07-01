import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import CdsProjectMutator from '../../cds-project-mutator.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  projectPath: string;
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
        option: '-p, --projectPath <projectPath>'
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

        if (!fs.existsSync(args.options.projectPath)) {
          return `Path ${args.options.projectPath} is not a valid path.`;
        }

        const existingSupportedProjects: string[] = this.getSupportedProjectFiles(args.options.projectPath);
        if (existingSupportedProjects.length === 0) {
          return `No supported project type found in path ${args.options.projectPath}.`;
        }

        if (existingSupportedProjects.length !== 1) {
          return `More than one supported project type found in path ${args.options.projectPath}.`;
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const referencedProjectFilePath: string = this.getSupportedProjectFiles(args.options.projectPath)[0];
      const relativeReferencedProjectFilePath: string = path.relative(process.cwd(), referencedProjectFilePath);
      const cdsProjectFilePath: string = this.getCdsProjectFile(process.cwd())[0];
      const cdsProjectFileContent: string = fs.readFileSync(cdsProjectFilePath, 'utf8');

      const cdsProjectMutator = new CdsProjectMutator(cdsProjectFileContent);
      cdsProjectMutator.addProjectReference(relativeReferencedProjectFilePath);

      fs.writeFileSync(cdsProjectFilePath, cdsProjectMutator.cdsProjectDocument as any);
    }
    catch (err: any) {
      throw err;
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

export default new PaSolutionReferenceAddCommand();