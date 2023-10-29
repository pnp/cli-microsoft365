import Command from "../Command.js";
import { CommandOptionInfo } from "./CommandOptionInfo.js";

export interface CommandInfo {
  aliases: string[] | undefined;
  command: Command;
  defaultProperties: string[] | undefined;
  description: string;
  file: string;
  help?: string;
  name: string;
  options: CommandOptionInfo[];
}