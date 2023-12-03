import { Cli, CommandOutput } from "./cli/Cli.js";

export async function executeCommand(commandName: string, options: any, listener?: {
  stdout: (message: any) => void,
  stderr: (message: any) => void,
}): Promise<CommandOutput> {
  const cli = Cli.getInstance();
  cli.loadAllCommandsInfo();
  await cli.loadCommandFromArgs(commandName.split(' '));
  if (!cli.commandToExecute) {
    return Promise.reject(`Command not found: ${commandName}`);
  }

  return Cli.executeCommandWithOutput(cli.commandToExecute.command, { options: options ?? {} }, listener);
}