export interface CommandOutput {
  error?: {
    message: string;
    code?: number;
  }
  stdout: string;
  stderr: string;
}

export declare function executeCommand(commandName: string, options: any, listener?: {
  stdout: (message: any) => void,
  stderr: (message: any) => void,
}): Promise<CommandOutput>;
