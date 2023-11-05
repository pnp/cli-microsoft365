import { CommandArgs } from "../Command";

export interface CommandOptionInfo {
  autocomplete: string[] | undefined;
  long?: string;
  name: string;
  required: boolean;
  short?: string;
  whenPrompted?: (optionName: string, args: CommandArgs) => Promise<any>;
  requiredWhen?: (args: CommandArgs) => boolean;
}