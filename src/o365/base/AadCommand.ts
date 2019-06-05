import Command from '../../Command';

export default abstract class AadCommand extends Command {
  protected get resource(): string {
    return 'https://graph.windows.net';
  }
}