import Command from '../../../Command';

export default abstract class PowerPlatformCommand extends Command {
  protected get graphResource(): string {
    return 'https://graph.microsoft.com';
  }

  protected get bapResource(): string {
    return 'https://api.bap.microsoft.com';
  }
}
