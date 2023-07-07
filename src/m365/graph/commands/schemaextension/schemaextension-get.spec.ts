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
const command: Command = require('./schemaextension-get');

describe(commands.SCHEMAEXTENSION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });
  it('gets schema extension', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        id: 'adatumisv_exo2'
      }
    });
    try {
      assert(loggerLogSpy.calledWith({
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
        "id": "adatumisv_exo2",
        "description": "sample description",
        "targetTypes": [
          "Message"
        ],
        "status": "Available",
        "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
        "properties": [
          {
            "name": "p1",
            "type": "String"
          },
          {
            "name": "p2",
            "type": "String"
          }
        ]
      }));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('gets schema extension(debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    });
    assert(loggerLogSpy.calledWith({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
      "id": "adatumisv_exo2",
      "description": "sample description",
      "targetTypes": [
        "Message"
      ],
      "status": "Available",
      "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
      "properties": [
        {
          "name": "p1",
          "type": "String"
        },
        {
          "name": "p2",
          "type": "String"
        }
      ]
    }));
  });
  it('handles error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
