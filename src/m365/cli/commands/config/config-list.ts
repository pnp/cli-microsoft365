import { z } from 'zod';
import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

export const options = z.strictObject({
  ...globalOptionsZod.shape
});

class CliConfigListCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_LIST;
  }

  public get description(): string {
    return 'Lists all self set CLI for Microsoft 365 configurations';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(cli.getConfig().all);
  }
}

export default new CliConfigListCommand();