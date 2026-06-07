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
  key: z.enum(settingNameValues).optional().alias('k')
});
type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliConfigResetCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_RESET;
  }

  public get description(): string {
    return 'Resets the specified CLI configuration option to its default value';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.key) {
      cli.getConfig().delete(args.options.key);
    }
    else {
      cli.getConfig().clear();
    }
  }
}

export default new CliConfigResetCommand();
