import auth from '../GraphAuth';
import { GraphItemsListCommand } from './GraphItemsListCommand';
import { GroupUser } from '../commands/o365group/GroupUser';

export abstract class GraphUsersListCommand<T> extends GraphItemsListCommand<GroupUser> {

  /* istanbul ignore next */
  constructor() {
    super();
  }

  protected getGroupUsers(cmd: CommandInstance, teamId: string): Promise<void> {
    return this.getOwners(cmd, teamId).then((): Promise<void> => this.getMembersAndGuests(cmd, teamId))
  }

  protected getOwners(cmd: CommandInstance, teamId: string): Promise<void> {
    const endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, true).then((): void => {
      // Currently there is a bug in the Microsoft Graph that returns Owners as
      // userType 'member'. We therefore update all returned user as owner  
      for (var i in this.items) {
        this.items[i].userType = 'Owner';
      }
    });
  }

  protected getMembersAndGuests(cmd: CommandInstance, teamId: string): Promise<void> {
    const endpoint: string = `${auth.service.resource}/v1.0/groups/${teamId}/members?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, false).then((): void => {
      this.items = this.filterDuplicateOwners();
    });
  }

  protected filterDuplicateOwners(): GroupUser[] {
    // Filter out duplicate added values for owners (as they are returned as members as well)
    return this.items.filter((groupUser, index, self) =>
      index === self.findIndex((t) => (
        t.id === groupUser.id && t.displayName === groupUser.displayName
      ))
    );
  }

}