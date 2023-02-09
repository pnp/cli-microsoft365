import { Logger } from '../../../../cli/Logger';
import { accessToken } from '../../../../utils/accessToken';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import auth from '../../../../Auth';

class PurviewRetentionEventListCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENT_LIST;
  }

  public get description(): string {
    return 'Get a list of retention events';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'eventTriggerDateTime'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);

      if (isAppOnlyAccessToken) {
        throw 'This command currently does not support app only permissions.';
      }

      if (this.verbose) {
        logger.logToStderr('Retrieving Purview retention events');
      }

      const items = await odata.getAllItems(`${this.resource}/beta/security/triggers/retentionEvents`);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewRetentionEventListCommand();