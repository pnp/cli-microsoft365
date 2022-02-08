import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  deleted?: boolean;
}

interface ExtendedGroup extends Group {
  groupType?: string;
}

class AadGroupListCommand extends GraphItemsListCommand<Group>   {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all groups defined in Azure Active Directory.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'groupType'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = args.options.deleted ? 'directory/deletedItems/microsoft.graph.group' : 'groups';

    this
      .getAllItems(`${this.resource}/v1.0/${endpoint}`, logger, true)
      .then((): void => {
        if (args.options.output === 'text') {
          this.items.forEach((group: ExtendedGroup) => {
            if (group.groupTypes && group.groupTypes.length > 0 && group.groupTypes[0] === 'Unified') {
              group.groupType = 'Microsoft 365';
            }
            else if (group.mailEnabled && group.securityEnabled) {
              group.groupType = 'Mail enabled security';
            }
            else if (group.securityEnabled) {
              group.groupType = 'Security';
            }
            else if (group.mailEnabled) {
              group.groupType = 'Distribution';
            }
          });
        }

        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-d, --deleted' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadGroupListCommand();