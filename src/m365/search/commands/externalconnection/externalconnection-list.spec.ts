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
const command: Command = require('./externalconnection-list');

describe(commands.EXTERNALCONNECTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnections: any = {
    value: [
      {
        "id": "contosohr",
        "name": "Contoso HR",
        "description": "Connection to index Contoso HR system",
        "state": "draft",
        "configuration": {
          "authorizedApps": [
            "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
          ],
          "authorizedAppIds": [
            "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
          ]
        }
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'state']);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
      }
    }), new CommandError('An error has occurred'));
  });

  it('retrieves list of external connections defined in the Microsoft Search', async () => {
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return externalConnections;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } } as any);
    assert(loggerLogSpy.calledWith(externalConnections.value));
  });
});
