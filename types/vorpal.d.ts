interface Vorpal {
  banner: (banner: string) => Vorpal;
  command: (command: string, description: string, autocomplete?: string[]) => VorpalCommand;
  _command: CurrentCommand;
  commands: CommandInfo[];
  delimiter: (delimiter: string) => Vorpal;
  description: (description: string) => Vorpal;
  exec: (command: string, callback?: () => void) => Promise<void>;
  find: (command: string) => VorpalCommand;
  isCommandArgKeyPairNormalized: boolean;
  log: (message: any) => void;
  on: (event: string, handler: (data?: any) => void) => Vorpal;
  parse: (argv: string[], options?: { use: string }) => Vorpal;
  pipe: (onStdout: (stdout: any) => any) => Vorpal;
  show: () => Vorpal;
  title: (title: string) => Vorpal;
  use: (extension: any) => Vorpal;
  version: (version: string) => Vorpal;
  chalk: any;
}

interface VorpalCommand {
  action: (action: (this: CommandInstance, args: any, callback: () => void) => void) => VorpalCommand;
  alias: (alias: string[]) => VorpalCommand;
  allowUnknownOptions: () => VorpalCommand;
  cancel: (handler: () => void) => VorpalCommand;
  help: (help: (args: any, cbOrLog: (message?: string) => void) => void) => VorpalCommand;
  helpInformation: () => string;
  hidden: () => VorpalCommand;
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

interface CommandInfo {
  options: CommandOption[];
  _args: CommandArg[];
  _aliases: string[];
  _name: string;
  _hidden: boolean;
}

interface CommandOption {
  autocomplete: string[];
  long: string;
  short: string;
}

interface CommandArg {
  name: string;
}