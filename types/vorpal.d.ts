interface Vorpal {
  command: (command: string, description: string, autocomplete?: string[]) => VorpalCommand;
  _command: CurrentCommand;
  delimiter: (delimiter: string) => Vorpal;
  exec: (command: string, callback?: () => void) => Promise<void>;
  find: (command: string) => VorpalCommand;
  on: (event: string, handler: (data?: any) => void) => Vorpal;
  parse: (argv: string[], options?: { use: string }) => Vorpal;
  pipe: (onStdout: (stdout: any) => any) => Vorpal;
  show: () => Vorpal;
  use: (extension: any) => Vorpal;
  chalk: any;
}

interface VorpalCommand {
  action: (action: (this: CommandInstance, args: any, callback: () => void) => void) => VorpalCommand;
  cancel: (handler: () => void) => VorpalCommand;
  help: (help: (args: any, log: (help: string) => void) => void) => VorpalCommand;
  helpInformation: () => string;
  option: (name: string, description?: string, autocomplete?: string[]) => VorpalCommand;
  types: (types: { string?: string[], boolean?: string[] }) => VorpalCommand;
  validate: (validator: (args: any) => boolean | string) => VorpalCommand;
}

interface CommandInstance {
  log: (message: any) => void;
  prompt: (object: any, callback: (result: any) => void) => void;
}

interface CurrentCommand {
  command: string;
  args: any;
}