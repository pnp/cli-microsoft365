import { AdministrativeUnit } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class AadAdministrativeUnitListCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of administrative units';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'visibility'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const results = await odata.getAllItems<AdministrativeUnit>(`${this.resource}/v1.0/directory/administrativeUnits`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadAdministrativeUnitListCommand();