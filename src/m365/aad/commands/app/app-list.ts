import { Application } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from "../../../../utils/odata";
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class AadAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of Azure AD app registrations';
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'id', 'displayName', "signInAudience"];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const results = await odata.getAllItems<Application>(`${this.resource}/v1.0/applications`);
      logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadAppListCommand();