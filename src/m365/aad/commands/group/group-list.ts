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
    return ['id', 'displayName', 'groupType'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groups`, logger, true)
      .then((): void => {
        if (args.options.output === 'text') {
          const allGroups: any = [];

          this.items.forEach(group => {
            if (group.groupTypes && group.groupTypes.length > 0 && group.groupTypes[0] === "Unified") {
              allGroups.push({id: group.id, displayName: group.displayName, groupType: "Microsoft 365"});
            }
            else if (group.mailEnabled && group.securityEnabled) {
              allGroups.push({id: group.id, displayName: group.displayName, groupType: "Mail enabled security"});
            }
            else if (group.securityEnabled) {
              allGroups.push({id: group.id, displayName: group.displayName, groupType: "Security"});
            }
            else if (group.mailEnabled) {
              allGroups.push({id: group.id, displayName: group.displayName, groupType: "Distribution"});
            }
          });

          logger.log(allGroups);
        }
        else {
          logger.log(this.items);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupListCommand();