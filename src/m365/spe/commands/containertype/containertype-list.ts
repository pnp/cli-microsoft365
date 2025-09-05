import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { odata } from '../../../../utils/odata.js';

class SpeContainerTypeListCommand extends GraphDelegatedCommand {

  public get name(): string {
    return commands.CONTAINERTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all container types';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'owningAppId'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of Container types...`);
      }

      const containerTypes = await odata.getAllItems<any>(`${this.resource}/beta/storage/fileStorage/containerTypes`);

      await logger.log(containerTypes);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerTypeListCommand();