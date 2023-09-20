import { Cli } from '../cli/Cli.js';
import { settingsNames } from '../settingsNames.js';

let inquirer: typeof import('inquirer') | undefined;

export const prompt = {
  /* c8 ignore next 10 */
  async forInput<T>(config: any, answers?: any): Promise<T> {
    if (!inquirer) {
      inquirer = await import('inquirer');
    }

    const cli = Cli.getInstance();
    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
    const prompt = inquirer.createPromptModule({ output: errorOutput === 'stderr' ? process.stderr : process.stdout });

    return await prompt(config, answers) as any;
  }
};