import Command from '../../Command.js';

export default abstract class O365MgmtCommand extends Command {
  protected get resource(): string {
    return 'https://manage.office.com';
  }
}