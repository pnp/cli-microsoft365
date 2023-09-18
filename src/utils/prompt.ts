let inquirerInput: typeof import('@inquirer/input') | undefined;
let inquirerConfirm: typeof import('@inquirer/confirm') | undefined;
let inquirerSelect: typeof import('@inquirer/select') | undefined;

export interface Choice<T> {
  name: string;
  value: T;
  description?: string;
}

export const prompt = {
  /* c8 ignore next 7 */
  async requestInput(message: string, defaultValue?: string): Promise<string> {
    if (!inquirerInput) {
      inquirerInput = await import('@inquirer/input');
    }

    return inquirerInput.default({ message, default: defaultValue }, { output: process.stderr });
  },

  /* c8 ignore next 7 */
  async requestConfirmation(message: string, defaultValue?: boolean): Promise<boolean> {
    if (!inquirerConfirm) {
      inquirerConfirm = await import('@inquirer/confirm');
    }

    return inquirerConfirm.default({ message, default: defaultValue }, { output: process.stderr });
  },

  /* c8 ignore next 7 */
  async requestSelection<T>(message: string, choices: Choice<T>[]): Promise<T> {
    if (!inquirerSelect) {
      inquirerSelect = await import('@inquirer/select');
    }

    return inquirerSelect.default({ message, choices }, { output: process.stderr });
  }
};