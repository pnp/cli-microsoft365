import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./groupsetting-get');

describe(commands.GROUPSETTING_GET, () => {
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUPSETTING_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified Group Setting', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert(loggerLogSpy.calledWith({
      "displayName": "Group Setting",
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
      "values": [
        {
          "name": "Name1",
          "value": "Value1"
        }
      ]
    }));
  });

  it('retrieves information about the specified Group Setting (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert(loggerLogSpy.calledWith({
      "displayName": "Group Setting",
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
      "values": [
        {
          "name": "Name1",
          "value": "Value1"
        }
      ]
    }));
  });

  it('correctly handles no group setting found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c843`) {
        throw {
          error: {
            "error": {
              "code": "Request_ResourceNotFound",
              "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.",
              "innerError": {
                "request-id": "7e192558-7438-46db-a4c9-5dca83d2ec96",
                "date": "2018-02-21T20:38:50"
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c843' } } as any),
      new CommandError(`Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
