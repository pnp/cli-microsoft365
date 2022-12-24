import Command from '../../Command';

export default abstract class O365MgmtCommand extends Command {
  protected get resource(): string {
    return 'https://manage.office.com';
  }
}