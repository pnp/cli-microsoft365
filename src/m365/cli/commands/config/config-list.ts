import { cli } from "../../../../cli/cli.js";
import { Logger } from "../../../../cli/Logger.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";

class CliConfigListCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_LIST;
  }

  public get description(): string {
    return 'List all self set CLI for Microsoft 365 configurations';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(cli.getConfig().all);
  }
}

export default new CliConfigListCommand();