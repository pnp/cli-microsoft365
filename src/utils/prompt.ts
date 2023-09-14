import { Separator } from '@inquirer/core';
import { settingsNames } from '../settingsNames.js';
import { Cli } from '../cli/Cli.js';

let inquirerInput: typeof import('@inquirer/input') | undefined;
let inquirerConfirm: typeof import('@inquirer/confirm') | undefined;
let inquirerSelect: typeof import('@inquirer/select') | undefined;

export interface Choice<T> {
  name: string;
  value: T;
  description?: string;
}

export interface InputConfig {
  message: string | Promise<string> | (() => Promise<string>);
  default?: string | undefined;
  transformer?: ((value: string, { isFinal }: {
    isFinal: boolean;
  }) => string) | undefined;
  validate?: ((value: string) => string | boolean | Promise<string | boolean>) | undefined;
}

export interface ConfirmationConfig {
  message: string | Promise<string> | (() => Promise<string>);
  default?: boolean | undefined;
  transformer?: ((value: boolean) => string) | undefined;
}

export interface SelectionConfig<Value> {
  message: string | Promise<string> | (() => Promise<string>);
  choices: readonly (Separator | Choice<Value>)[];
  pageSize?: number | undefined;
}

export const prompt = {
  /* c8 ignore next 10 */
  async forInput(config: InputConfig): Promise<string> {
    if (!inquirerInput) {
      inquirerInput = await import('@inquirer/input');
    }

    const cli = Cli.getInstance();
    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');

    return inquirerInput.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  },

  /* c8 ignore next 10 */
  async forConfirmation(config: ConfirmationConfig): Promise<boolean> {
    if (!inquirerConfirm) {
      inquirerConfirm = await import('@inquirer/confirm');
    }

    const cli = Cli.getInstance();
    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');

    return inquirerConfirm.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  },

  /* c8 ignore next 10 */
  async forSelection<T>(config: SelectionConfig<T>): Promise<T> {
    if (!inquirerSelect) {
      inquirerSelect = await import('@inquirer/select');
    }

    const cli = Cli.getInstance();
    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');

    return inquirerSelect.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  }
};