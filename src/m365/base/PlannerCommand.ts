import Command from '../../Command.js';

export default abstract class PlannerCommand extends Command {
  protected get resource(): string {
    return 'https://tasks.office.com';
  }
}
