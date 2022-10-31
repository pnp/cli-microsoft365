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

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const url = `https://graph.microsoft.com/v1.0/applications`;

      logger.log(await odata.getAllItems<Application>(url));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'id', 'displayName', "signInAudience"];
  }
}

module.exports = new AadAppListCommand();