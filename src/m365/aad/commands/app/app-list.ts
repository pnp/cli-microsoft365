import { Application } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from "../../../../utils/odata";
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

export interface CommandArgs {
  options: GlobalOptions;
}

class AadAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Gets a list of Azure AD app registrations';
  }

  constructor() {
    super();
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const appList = await this.getAppList();
      logger.log(appList);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getAppList(): Promise<Application[]> {
    const select = ["appId","id","displayName","signInAudience"];
    const url = `https://graph.microsoft.com/v1.0/applications?$select=${select.join(",")}`;

    return odata.getAllItems<Application>(url);
  }

}

module.exports = new AadAppListCommand();