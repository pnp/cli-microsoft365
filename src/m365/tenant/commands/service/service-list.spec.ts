import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import Command, { CommandError, CommandOption } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./service-list');

describe(commands.TENANT_SERVICE_LIST, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  const textOutput = [
    {
      Id: "Bookings",
      DisplayName: "Microsoft Bookings"
    },
    {
      Id: "DynamicsCRM",
      DisplayName: "Dynamics 365"
    }
  ];

  const jsonOutput = {
    "value": [
      {
        "Id": "Bookings",
        "DisplayName": "Microsoft Bookings",
        "Features": [
          {
            "DisplayName": "Microsoft Bookings",
            "Name": "MicrosoftBookings"
          }
        ]
      },
      {
        "Id": "DynamicsCRM",
        "DisplayName": "Dynamics 365",
        "Features": [
          {
            "DisplayName": "Sign In",
            "Name": "signin"
          },
          {
            "DisplayName": "Sign up and administration",
            "Name": "admin"
          },
          {
            "DisplayName": "Organization access",
            "Name": "orgaccess"
          },
          {
            "DisplayName": "Organization performance",
            "Name": "orgperf"
          },
          {
            "DisplayName": "Components/Features",
            "Name": "crmcomponents"
          }
        ]
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(Utils, 'getTenantIdFromAccessToken').callsFake(() => {
      return '31537af4-6d77-4bb9-a681-d2394888ea26';
    });

    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      Utils.getTenantIdFromAccessToken
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TENANT_SERVICE_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('handles promise error while getting services available in Microsoft 365', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('ServiceComms/Services') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {

      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the services available in Microsoft 365', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('ServiceComms/Services') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the services available in Microsoft 365 (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('ServiceComms/Services') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(jsonOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the services available in Microsoft 365 as text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('ServiceComms/Services') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the services available in Microsoft 365 as text (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('ServiceComms/Services') > -1) {
        return Promise.resolve(jsonOutput);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(textOutput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 