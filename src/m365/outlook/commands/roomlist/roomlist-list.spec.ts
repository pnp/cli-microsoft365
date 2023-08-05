import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './roomlist-list.js';

describe(commands.ROOMLIST_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const jsonOutput = {
    "value": [
      {
        "id": "DC404124-302A-92AA-F98D-7B4DEB0C1705",
        "displayName": "Building 1",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geocoordinates": null,
        "phone": null,
        "emailAddress": "bldg1@contoso.com"
      },
      {
        "id": "DC404124-302A-92AA-F98D-7B4DEB0C1706",
        "displayName": "Building 2",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geocoordinates": null,
        "phone": null,
        "emailAddress": "bldg2@contoso.com"
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
    assert.strictEqual(command.name, commands.ROOMLIST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'phone', 'emailAddress']);
  });

  it('lists all available roomlist in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/microsoft.graph.roomlist`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: { force: true } }), new CommandError(errorMessage));
  });
});
