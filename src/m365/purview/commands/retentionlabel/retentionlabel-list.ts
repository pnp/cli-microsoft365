import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import { accessToken } from '../../../../utils/accessToken';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class PurviewRetentionLabelListCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_LIST;
  }

  public get description(): string {
    return 'Get a list of retention labels';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'isInUse'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
    }

    try {
      const items = await odata.getAllItems(`${this.resource}/beta/security/labels/retentionLabels`);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewRetentionLabelListCommand();