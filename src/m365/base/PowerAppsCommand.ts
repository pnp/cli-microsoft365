import Command from '../../Command';

export default abstract class PowerAppsCommand extends Command {
  protected get resource(): string {
    return 'https://api.powerapps.com';
  }
}