import { CommandOutput, cli } from "./cli/cli.js";

export async function executeCommand(commandName: string, options: any, listener?: {
  stdout: (message: any) => void,
  stderr: (message: any) => void,
}): Promise<CommandOutput> {
  cli.loadAllCommandsInfo();
  await cli.loadCommandFromArgs(commandName.split(' '));
  if (!cli.commandToExecute) {
    throw `Command not found: ${commandName}`;
  }

  return cli.executeCommandWithOutput(cli.commandToExecute.command, { options: options ?? {} }, listener);
}