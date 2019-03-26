import Command from '../../Command';

export default abstract class AzmgmtCommand extends Command {
  protected get resource(): string {
    return 'https://management.azure.com/';
  }
}