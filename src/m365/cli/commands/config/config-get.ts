import { z } from 'zod';
import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import { settingsNames } from "../../../../settingsNames.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

const settingNameValues = Object.getOwnPropertyNames(settingsNames) as [string, ...string[]];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  key: z.enum(settingNameValues).alias('k')
});
type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliConfigGetCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_GET;
  }

  public get description(): string {
    return 'Gets value of a CLI for Microsoft 365 configuration option';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await logger.log(cli.getConfig().get(args.options.key));
  }
}

export default new CliConfigGetCommand();