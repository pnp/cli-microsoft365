import { Cli } from "../../../../cli/Cli";
import { Logger } from "../../../../cli/Logger";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";

class CliConfigListCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_LIST;
  }

  public get description(): string {
    return 'Gets value of a CLI for Microsoft 365 configuration option';
  }

  constructor() {
    super();
  }

  public async commandAction(logger: Logger): Promise<void> {
    logger.log(Cli.getInstance().config.all);
  }
}

module.exports = new CliConfigListCommand();