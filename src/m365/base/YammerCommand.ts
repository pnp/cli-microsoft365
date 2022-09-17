import Command, { CommandError } from "../../Command";

export default abstract class YammerCommand extends Command {
  protected get resource(): string {
    return 'https://www.yammer.com/api';
  }

  protected handleRejectedODataJsonPromise(response: any): void {
    if (response.statusCode === 404) {
      throw new CommandError("Not found (404)");
    }
    else if (response.error && response.error.base) {
      throw new CommandError(response.error.base);
    }
    else {
      throw new CommandError(response);
    }
  }
}