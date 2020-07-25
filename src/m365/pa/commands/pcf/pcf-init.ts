import * as fs from "fs";
import * as path from 'path';
import { v4 } from 'uuid';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandAction,
  CommandError
} from '../../../../Command';
import Utils from '../../../../Utils';
import TemplateInstantiator from "../../template-instantiator";
import { PcfInitVariables } from "./pcf-init/pcf-init-variables";
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
class PaPcfInitCommand extends Command {
  public get name(): string {
    return commands.PCF_INIT;
  }

  public get description(): string {
    return 'Creates new PowerApps component framework project';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.template = args.options.template;
    return telemetryProps;
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
        cmd.log(`name: ${args.options.name}`);
        cmd.log(`namespace: ${args.options.namespace}`);
        cmd.log(`template: ${args.options.template}`);
        cmd.log(`pcfTemplatePath: ${pcfTemplatePath}`);
        cmd.log(`pcfComponentTemplatePath: ${pcfComponentTemplatePath}`);
        cmd.log(`workingDirectory: ${workingDirectory}`);
        cmd.log(`workingDirectoryName: ${workingDirectoryName}`);
        cmd.log(`componentDirectory: ${componentDirectory}`);
      }

      TemplateInstantiator.instantiate(cmd, pcfTemplatePath, workingDirectory, false, variables, this.verbose);
      TemplateInstantiator.instantiate(cmd, pcfComponentTemplatePath, componentDirectory, true, variables, this.verbose);

      if (this.verbose) {
        cmd.log(` `);
      }

      cmd.log(chalk.green(`The PowerApps component framework project was successfully created in '${workingDirectory}'.`));
      cmd.log(`Be sure to run '${chalk.grey('npm install')}' in this directory to install project dependencies.`);

      cb();
    }
    catch (err) {
      cb(new CommandError(err));
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--namespace <namespace>',
        description: 'The namespace for the component.'
      },
      {
        option: '--name <name>',
        description: 'The name for the component.',
      },
      {
        option: '--template <template>',
        description: 'Choose a template for the component. Field|Dataset.',
        autocomplete: ['Field', 'Dataset']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      if (fs.readdirSync(process.cwd()).some(fn => path.extname(fn).toLowerCase().endsWith('proj'))) {
        return 'PowerApps component framework project creation failed. The current directory already contains a project. Please create a new directory and retry the operation.';
      }

      const workingDirectoryName: string = path.basename(process.cwd());
      if (!Utils.isValidFileName(workingDirectoryName))
      {
          return `Empty or invalid project name '${workingDirectoryName}'`;
      }

      if (args.options.name) {
        if (!/^(?!\d)[a-zA-Z0-9]+$/i.test(args.options.name)) {
          return `Value of 'name' is invalid. Only characters within the ranges [A - Z], [a - z] or [0 - 9] are allowed. The first character may not be a number.`;
        }

        if (Utils.isJavascriptReservedWord(args.options.name)) {
          return `The value '${args.options.name}' passed for 'name' is a reserved word.`;
        }
      }
      else {
        return 'Missing required option name.';
      }

      if (args.options.namespace) {
        if (!/^(?!\.|\d)(?!.*\.$)(?!.*?\.\d)(?!.*?\.\.)[a-zA-Z0-9.]+$/i.test(args.options.namespace)) {
          return `Value of 'namespace' is invalid. Only characters within the ranges [A - Z], [a - z], [0 - 9], or '.' are allowed. The first and last character may not be the '.' character. Consecutive '.' characters are not allowed. Numbers are not allowed as the first character or immediately after a period.`;
        }

        if (Utils.isJavascriptReservedWord(args.options.namespace)) {
          return `The value '${args.options.namespace}' passed for 'namespace' is or includes a reserved word.`;
        }
      }
      else {
        return 'Missing required option namespace.';
      }

      if (args.options.namespace && args.options.name && (args.options.namespace + args.options.name).length > 75) {
        return `The total length of values for 'name' and 'namespace' cannot exceed 75. Length of 'name' is ${args.options.name.length}, length of 'namespace' is ${args.options.namespace.length}.`;
      }

      if (args.options.template) {
        const testTemplate: string = args.options.template.toLowerCase();
        if (!(testTemplate === 'field' || testTemplate === 'dataset')) {
          return `Template must be either 'Field' or 'Dataset', but '${args.options.template}' was provided.`;
        }
      }
      else {
        return 'Missing required option template.';
      }

      return true;
    };
  }
}

module.exports = new PaPcfInitCommand();