import { Application } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from "../../../../utils/odata.js";
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

class EntraAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of Entra app registrations';
  }

  public alias(): string[] | undefined {
    return [aadCommands.APP_LIST, commands.APPREGISTRATION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['appId', 'id', 'displayName', "signInAudience"];
  }

  public async commandAction(logger: Logger): Promise<void> {
    this.showDeprecationWarning(logger, aadCommands.APP_LIST, commands.APP_LIST);

    try {
      const results = await odata.getAllItems<Application>(`${this.resource}/v1.0/applications`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAppListCommand();