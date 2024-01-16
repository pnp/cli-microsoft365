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
import command from './room-list.js';

describe(commands.ROOM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const jsonOutput = {
    "value": [
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78989EB1",
        "emailAddress": "cf100@contoso.com",
        "displayName": "Conf Room 100",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "100",
        "capacity": 50,
        "building": "1",
        "floorNumber": 1,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "bean bags"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      },
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78970B97",
        "emailAddress": "cf200@contoso.com",
        "displayName": "Conf Room 200",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "200",
        "capacity": 40,
        "building": "2",
        "floorNumber": 2,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "benches",
          "nice view"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      }
    ]
  };
  const jsonOutputFilter = {
    "value": [
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78970B97",
        "emailAddress": "cf200@contoso.com",
        "displayName": "Conf Room 200",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "200",
        "capacity": 40,
        "building": "2",
        "floorNumber": 2,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "benches",
          "nice view"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      }
    ]
  };
  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.ROOM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'phone', 'emailAddress']);
  });

  it('lists all available rooms in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/microsoft.graph.room`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all available rooms filter by roomlistEmail in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/bldg2@contoso.com/microsoft.graph.roomlist/rooms`) {
        return jsonOutputFilter;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roomlistEmail: "bldg2@contoso.com" } });
    assert(loggerLogSpy.calledWith(
      jsonOutputFilter.value
    ));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: { force: true } }), new CommandError(errorMessage));
  });
});
