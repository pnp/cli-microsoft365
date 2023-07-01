import Command from "../Command.js";
import { CommandOptionInfo } from "./CommandOptionInfo.js";

export interface CommandInfo {
  aliases: string[] | undefined;
  command: Command;
  defaultProperties: string[] | undefined;
  name: string;
  options: CommandOptionInfo[];
}