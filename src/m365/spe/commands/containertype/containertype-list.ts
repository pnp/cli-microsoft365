import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ContainerTypeProperties, spo } from '../../../../utils/spo.js';

class SpeContainertypeListCommand extends SpoCommand {

  public get name(): string {
    return commands.CONTAINERTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all Container Types';
  }

  public defaultProperties(): string[] | undefined {
    return ['ContainerTypeId', 'DisplayName', 'OwningAppId'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of Container types...`);
      }

      const allContainerTypes: ContainerTypeProperties[] = await spo.getAllContainerTypes(spoAdminUrl, logger, this.debug);
      await logger.log(allContainerTypes);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpeContainertypeListCommand();