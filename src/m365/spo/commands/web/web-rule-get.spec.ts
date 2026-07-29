import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './web-rule-get.js';

describe(commands.WEB_RULE_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const webUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const alertId = '39d9e102-9e8f-4e74-8f17-84a92f972fcf';
  const alertResponse = {
    AlertFrequency: 0,
    AlertTemplateName: 'SPAlertTemplateType.DocumentLibrary',
    AlertType: 0,
    AlwaysNotify: false,
    DeliveryChannels: 1,
    EventType: -1,
    Filter: '',
    ID: alertId,
    Properties: [
      {
        Key: 'webUrl',
        Value: 'https://contoso.sharepoint.com',
        ValueType: 'Edm.String'
      }
    ],
    Status: 0,
    Title: 'Marketing documents',
    UserId: 8,
    List: {
      Id: '7cbb4c8d-8e4d-4d2e-9c6f-3f1d8b2e6a0e',
      Title: 'Documents',
      RootFolder: {
        ServerRelativeUrl: '/sites/marketing/Shared Documents'
      }
    },
    User: {
      Id: 8,
      UserPrincipalName: 'jane.doe@contoso.com'
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEB_RULE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', id: alertId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if alertId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl, id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when valid webUrl and alertId are provided', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl, id: alertId });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves a rule by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/Alerts/GetById('${formatting.encodeQueryParameter(alertId)}')?$expand=List,User,List/Rootfolder&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl`) {
        return alertResponse;
      }

      throw new Error(`Invalid request: ${opts.url}`);
    });

    await command.action(logger, { options: { webUrl, id: alertId, verbose: true } });
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('logs verbose output to stderr', async () => {
    sinon.stub(request, 'get').resolves(alertResponse);

    await command.action(logger, { options: { webUrl, id: alertId, verbose: true } });
    assert(loggerLogToStderrSpy.calledWith(`Retrieving rule with id '${alertId}' from site '${webUrl}'...`));
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2146232832, Microsoft.SharePoint.SPException',
          message: {
            value: 'The alert you are trying to access does not exist or has just been deleted.'
          }
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { webUrl, id: alertId } }),
      new CommandError(error.error['odata.error'].message.value));
  });
});
