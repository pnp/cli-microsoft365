import Command from '../../Command.js';

export default abstract class PowerBICommand extends Command {
  protected get resource(): string {
    return 'https://api.powerbi.com';
  }
}
