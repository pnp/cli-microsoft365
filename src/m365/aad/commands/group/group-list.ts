import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Group } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: GlobalOptions;
}

class AadGroupListCommand extends GraphItemsListCommand<Group>   {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all groups defined in Azure Active Directory.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'visibility'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groups`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupListCommand();