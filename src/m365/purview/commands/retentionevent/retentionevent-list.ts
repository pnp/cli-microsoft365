import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod.strict();

class PurviewRetentionEventListCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENT_LIST;
  }

  public get description(): string {
    return 'Get a list of retention events';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'eventTriggerDateTime'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving Purview retention events');
      }

      const items = await odata.getAllItems(`${this.resource}/v1.0/security/triggers/retentionEvents`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionEventListCommand();