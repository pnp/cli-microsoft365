export interface CommandOutput {
  error?: {
    message: string;
    code?: number;
  }
  stdout: string;
  stderr: string;
}

export declare function executeCommand(commandName: string, options: any): Promise<CommandOutput>;
export declare function on(eventName: string, listener: (...args: any[]) => void): void;
