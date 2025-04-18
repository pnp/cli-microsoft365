import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spe } from '../../../../utils/spe.js';
import { spo } from '../../../../utils/spo.js';

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

      const allContainerTypes = await spe.getAllContainerTypes(spoAdminUrl);

      // The following conversion is done in order not to make breaking changes
      const result = allContainerTypes.map(ct => ({
        _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
        ...ct,
        AzureSubscriptionId: `/Guid(${ct.AzureSubscriptionId})/`,
        ContainerTypeId: `/Guid(${ct.ContainerTypeId})/`,
        OwningAppId: `/Guid(${ct.OwningAppId})/`,
        OwningTenantId: `/Guid(${ct.OwningTenantId})/`
      }));

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpeContainertypeListCommand();