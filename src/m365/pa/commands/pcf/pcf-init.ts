import chalk from 'chalk';
import fs from 'fs';
import path from 'path';
import url from 'url';
import { v4 } from 'uuid';
import { Logger } from "../../../../cli/Logger.js";
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from '../../commands.js';
import TemplateInstantiator from "../../template-instantiator.js";
import { PcfInitVariables } from "./pcf-init/pcf-init-variables.js";

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  namespace: string;
  name: string;
  template: string;
}

/*
 * Logic extracted from bolt.module.pcf.dll
 * Version: 1.0.6
 * Class: bolt.module.pcf.PcfInitVerb
 */
class PaPcfInitCommand extends AnonymousCommand {
  public get name(): string {
    return commands.PCF_INIT;
  }

  public get description(): string {
    return 'Creates new PowerApps component framework project';
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
        template: args.options.template
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--namespace <namespace>'
      },
      {
        option: '--name <name>'
      },
      {
        option: '--template <template>',
        autocomplete: ['Field', 'Dataset']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (fs.readdirSync(process.cwd()).some(fn => path.extname(fn).toLowerCase().endsWith('proj'))) {
          return 'PowerApps component framework project creation failed. The current directory already contains a project. Please create a new directory and retry the operation.';
        }

        const workingDirectoryName: string = path.basename(process.cwd());
        if (!validation.isValidFileName(workingDirectoryName)) {
          return `Empty or invalid project name '${workingDirectoryName}'`;
        }

        if (!/^(?!\d)[a-zA-Z0-9]+$/i.test(args.options.name)) {
          return `Value of 'name' is invalid. Only characters within the ranges [A - Z], [a - z] or [0 - 9] are allowed. The first character may not be a number.`;
        }

        if (validation.isJavaScriptReservedWord(args.options.name)) {
          return `The value '${args.options.name}' passed for 'name' is a reserved word.`;
        }

        if (!/^(?!\.|\d)(?!.*\.$)(?!.*?\.\d)(?!.*?\.\.)[a-zA-Z0-9.]+$/i.test(args.options.namespace)) {
          return `Value of 'namespace' is invalid. Only characters within the ranges [A - Z], [a - z], [0 - 9], or '.' are allowed. The first and last character may not be the '.' character. Consecutive '.' characters are not allowed. Numbers are not allowed as the first character or immediately after a period.`;
        }

        if (validation.isJavaScriptReservedWord(args.options.namespace)) {
          return `The value '${args.options.namespace}' passed for 'namespace' is or includes a reserved word.`;
        }

        if (args.options.namespace && args.options.name && (args.options.namespace + args.options.name).length > 75) {
          return `The total length of values for 'name' and 'namespace' cannot exceed 75. Length of 'name' is ${args.options.name.length}, length of 'namespace' is ${args.options.namespace.length}.`;
        }

        const testTemplate: string = args.options.template.toLowerCase();
        if (!(testTemplate === 'field' || testTemplate === 'dataset')) {
          return `Template must be either 'Field' or 'Dataset', but '${args.options.template}' was provided.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const pcfTemplatePath: string = path.join(__dirname, 'pcf-init', 'assets');
      const pcfComponentTemplatePath: string = path.join(pcfTemplatePath, 'control', `${args.options.template.toLowerCase()}-template`);
      const workingDirectory: string = process.cwd();
      const workingDirectoryName: string = path.basename(workingDirectory);
      const componentDirectory: string = path.join(workingDirectory, args.options.name);
      const variables: PcfInitVariables = {
        "$namespaceplaceholder$": args.options.namespace,
        "$controlnameplaceholder$": args.options.name,
        "$pcfProjectName$": workingDirectoryName,
        "pcfprojecttype": workingDirectoryName,
        "$pcfProjectGuid$": v4()
      };

      if (this.verbose) {
        await logger.logToStderr(`name: ${args.options.name}`);
        await logger.logToStderr(`namespace: ${args.options.namespace}`);
        await logger.logToStderr(`template: ${args.options.template}`);
        await logger.logToStderr(`pcfTemplatePath: ${pcfTemplatePath}`);
        await logger.logToStderr(`pcfComponentTemplatePath: ${pcfComponentTemplatePath}`);
        await logger.logToStderr(`workingDirectory: ${workingDirectory}`);
        await logger.logToStderr(`workingDirectoryName: ${workingDirectoryName}`);
        await logger.logToStderr(`componentDirectory: ${componentDirectory}`);
      }

      TemplateInstantiator.instantiate(logger, pcfTemplatePath, workingDirectory, false, variables, this.verbose);
      TemplateInstantiator.instantiate(logger, pcfComponentTemplatePath, componentDirectory, true, variables, this.verbose);

      if (this.verbose) {
        await logger.logToStderr(` `);
      }

      await logger.log(chalk.green(`The PowerApps component framework project was successfully created in '${workingDirectory}'.`));
      await logger.log(`Be sure to run '${chalk.grey('npm install')}' in this directory to install project dependencies.`);
    }
    catch (err: any) {
      throw err;
    }
  }
}

export default new PaPcfInitCommand();