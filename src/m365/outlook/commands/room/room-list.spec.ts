import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./room-list');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ROOM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'phone', 'emailAddress']);
  });

  it('lists all available rooms in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/microsoft.graph.room`) {
        return Promise.resolve(
          jsonOutput
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutput.value
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all available rooms filter by roomlistEmail in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/bldg2@contoso.com/microsoft.graph.roomlist/rooms`) {
        return Promise.resolve(
          jsonOutputFilter
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true, roomlistEmail: "bldg2@contoso.com" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutputFilter.value
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
