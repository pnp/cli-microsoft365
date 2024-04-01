import { Separator } from '@inquirer/core';
import { settingsNames } from '../settingsNames.js';
import { cli } from '../cli/cli.js';

let inquirerInput: typeof import('@inquirer/input') | undefined;
let inquirerConfirm: typeof import('@inquirer/confirm') | undefined;
let inquirerSelect: typeof import('@inquirer/select') | undefined;

export interface Choice<T> {
  name: string;
  value: T;
  description?: string;
}

export interface InputConfig {
  message: string;
  default?: string;
  transformer?: (value: string, { isFinal }: {
    isFinal: boolean;
  }) => string;
  validate?: (value: string) => string | boolean | Promise<string | boolean>;
  theme?: PartialDeep<Theme>;
}

export interface ConfirmationConfig {
  message: string;
  default?: boolean;
  transformer?: ((value: boolean) => string);
  theme?: PartialDeep<Theme>;
}

export interface SelectionConfig<Value> {
  message: string;
  choices: readonly (Separator | Choice<Value>)[];
  pageSize?: number;
  loop?: boolean;
  default?: unknown;
  theme?: PartialDeep<SelectTheme>;
}

interface Theme {
  prefix: string;
  spinner: {
    interval: number;
    frames: string[];
  };
  style: {
    answer: (text: string) => string;
    message: (text: string) => string;
    error: (text: string) => string;
    defaultAnswer: (text: string) => string;
    help: (text: string) => string;
    highlight: (text: string) => string;
    key: (text: string) => string;
  };
};

interface SelectTheme extends Theme {
  icon: { cursor: string };
  style: {
    answer: (text: string) => string;
    message: (text: string) => string;
    error: (text: string) => string;
    defaultAnswer: (text: string) => string;
    help: (text: string) => string;
    highlight: (text: string) => string;
    key: (text: string) => string;
    disabled: (text: string) => string;
  };
};

type PartialDeep<T> = T extends object
  ? { [P in keyof T]?: PartialDeep<T[P]>; }
  : T;

export const prompt = {
  /* c8 ignore next 9 */
  async forInput(config: InputConfig): Promise<string> {
    if (!inquirerInput) {
      inquirerInput = await import('@inquirer/input');
    }

    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');

    return inquirerInput.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  },

  /* c8 ignore next 9 */
  async forConfirmation(config: ConfirmationConfig): Promise<boolean> {
    if (!inquirerConfirm) {
      inquirerConfirm = await import('@inquirer/confirm');
    }

    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');

    return inquirerConfirm.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  },

  /* c8 ignore next 14 */
  async forSelection<T>(config: SelectionConfig<T>): Promise<T> {
    if (!inquirerSelect) {
      inquirerSelect = await import('@inquirer/select');
    }

    const errorOutput: string = cli.getSettingWithDefaultValue(settingsNames.errorOutput, 'stderr');
    const promptPageSizeCap: number = cli.getSettingWithDefaultValue(settingsNames.promptListPageSize, 7);

    if (!config.pageSize) {
      config.pageSize = Math.min(config.choices.length, promptPageSizeCap);
    }

    return inquirerSelect.default(config, { output: errorOutput === 'stderr' ? process.stderr : process.stdout });
  }
};