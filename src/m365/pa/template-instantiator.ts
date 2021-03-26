import * as fs from "fs";
import * as path from "path";
import { Logger } from "../../cli";
import { PcfInitVariables } from "./commands/pcf/pcf-init/pcf-init-variables";
import { SolutionInitVariables } from "./commands/solution/solution-init/solution-init-variables";

/*
 * Logic extracted from bolt.cli.dll
 * Version: 1.0.6
 * Class: bolt.cli.TemplateInstantiator
 */
export default class TemplateInstantiator {
  public static instantiate(logger: Logger, sourcePath: string, destinationPath: string, recursive: boolean, variables: PcfInitVariables | SolutionInitVariables, verbose: boolean): void {
    TemplateInstantiator.mkdirSyncIfNotExists(logger, destinationPath, verbose);

    this.getFiles(sourcePath, recursive).forEach(file => {
      const filePath = path.relative(sourcePath, path.dirname(file));
      const destinationFilePath = path.join(destinationPath, filePath);

      TemplateInstantiator.mkdirSyncIfNotExists(logger, destinationFilePath, verbose);

      this.instantiateTemplate(file, destinationFilePath, variables);
    });
  }

  public static mkdirSyncIfNotExists(logger: Logger, destinationPath: string, verbose: boolean): void {
    if (!fs.existsSync(destinationPath)) {
      if (verbose) {
        logger.logToStderr(`Create directory: ${destinationPath}`);
      }
      fs.mkdirSync(destinationPath);
    }
  }

  private static instantiateTemplate(templatePath: string, destinationPath: string, variables: PcfInitVariables | SolutionInitVariables) {
    let isTemplateFile: boolean = false;
    let fileName: string = path.basename(templatePath);

    if (fileName.toLowerCase().startsWith('template_')) {
      isTemplateFile = true;
      fileName = fileName.substring('template_'.length, fileName.length);
    }

    for (const variable in variables) {
      fileName = fileName.replace(variable, variables[variable]);
    }

    const destinationFilePath: string = path.join(destinationPath, fileName);

    if (!isTemplateFile) {
      fs.copyFileSync(templatePath, destinationFilePath);
    }
    else {
      let fileContent = fs.readFileSync(templatePath, 'utf8');

      for (const variable in variables) {
        fileContent = fileContent.replace(new RegExp(variable.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), 'g'), variables[variable]);
      }

      fs.writeFileSync(destinationFilePath, fileContent, 'utf8');
    }
  }

  private static getFiles(folderPath: string, recursive: boolean): string[] {
    const entryPaths = fs.readdirSync(folderPath).map(entry => path.join(folderPath, entry));
    const filePaths = entryPaths.filter(entryPath => fs.statSync(entryPath).isFile());
    const dirPaths = entryPaths.filter(entryPath => !filePaths.includes(entryPath));
    const dirFiles = recursive ? dirPaths.reduce((prev, curr) => prev.concat(this.getFiles(curr, recursive)), [] as string[]) : [];
    return [...filePaths, ...dirFiles];
  }
}