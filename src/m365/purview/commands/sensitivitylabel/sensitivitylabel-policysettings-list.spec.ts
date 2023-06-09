import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./sensitivitylabel-policysettings-list');

describe(commands.SENSITIVITYLABEL_POLICYSETTINGS_LIST, () => {
  const userId = '59f80e08-24b1-41f8-8586-16765fd830d3';
  const userName = 'john.doe@contoso.com';

  const sensitivityLabelPolicySettingsListResponse = {
    "id": "71F139249895C2F6DC861031DAC47E0C2C37C6595582D4248CC77FD7293681B5DE348BC71AEB44068CB397DB021CADB4",
    "moreInfoUrl": "https://docs.microsoft.com/en-us/microsoft-365/compliance/get-started-with-sensitivity-labels?view=o365-worldwide#end-user-documentation-for-sensitivity-labels",
    "isMandatory": true,
    "isDowngradeJustificationRequired": true,
    "defaultLabelId": "022bb90d-0cda-491d-b861-d195b14532dc"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SENSITIVITYLABEL_POLICYSETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userId defined', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userName defined', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves list of policy settings for a sensitivity label that the current logged in user has access to', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/me/security/informationProtection/labelPolicySettings`) {
        return sensitivityLabelPolicySettingsListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(sensitivityLabelPolicySettingsListResponse));
  });

  it('retrieves list of policy settings for a sensitivity label that the specific user has access to by its Id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/security/informationProtection/labelPolicySettings`) {
        return sensitivityLabelPolicySettingsListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });
    assert(loggerLogSpy.calledWith(sensitivityLabelPolicySettingsListResponse));
  });

  it('retrieves list of policy settings for a sensitivity label that the specific user has access to by its UPN', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/security/informationProtection/labelPolicySettings`) {
        return sensitivityLabelPolicySettingsListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(sensitivityLabelPolicySettingsListResponse));
  });

  it('throws an error when using application permissions and no option is specified', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError(`Specify at least 'userId' or 'userName' when using application permissions.`));
  });

  it('handles error when list of policy settings for a sensitivity label is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/security/informationProtection/labelPolicySettings`) {
        throw `Error: The resource could not be found.`;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(`Error: The resource could not be found.`));
  });
});