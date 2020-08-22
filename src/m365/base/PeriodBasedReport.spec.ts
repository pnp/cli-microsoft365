import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
import * as assert from 'assert';
import Utils from '../../Utils';
import request from '../../request';
import * as fs from 'fs';
import PeriodBasedReport from './PeriodBasedReport';
import { CommandValidate, CommandError } from '../../Command';

class MockCommand extends PeriodBasedReport {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public get usageEndpoint(): string {
    return 'MockEndPoint';
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('PeriodBasedReport', () => {
  const mockCommand = new MockCommand();
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  //let cmdInstanceLogSpy: sinon.SinonSpy;
  //let writeFileSyncFake = () => { };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: mockCommand.name
      },
      action: mockCommand.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    //cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (mockCommand as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      fs.writeFileSync
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(mockCommand.name.startsWith('mock'), true);
  });

  it('has a description', () => {
    assert.notEqual(mockCommand.description, null);
  });

  it('fails validation if period option is not passed', () => {
    const actual = (mockCommand.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation on invalid period', () => {
    const actual = (mockCommand.validate() as CommandValidate)({ options: { period: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation on valid \'D7\' period', () => {
    const actual = (mockCommand.validate() as CommandValidate)({
      options: {
        period: 'D7'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'D30\' period', () => {
    const actual = (mockCommand.validate() as CommandValidate)({
      options: {
        period: 'D30'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'D90\' period', () => {
    const actual = (mockCommand.validate() as CommandValidate)({
      options: {
        period: 'D90'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'180\' period', () => {
    const actual = (mockCommand.validate() as CommandValidate)({
      options: {
        period: 'D90'
      }
    });
    assert.equal(actual, true);
  });


  it('get unique device type in teams and export it in a period', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('produce export using period format and Teams unique device type output in txt', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    //const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        //assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('produce export using period format and Teams unique device type output in json', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    //const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7', output: 'json' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        //assert.equal(fileStub.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('produce export using period format and Teams unique users output in txt', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });
    //const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7',  output: 'text' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        //assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('produce export using period format and Teams unique users output in json', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `);
      }

      return Promise.reject('Invalid request');
    });
    //const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        //assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('produce export using period format and Teams output in json', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return Promise.resolve(`Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period\n2019-08-28,0,0,0,0,0,0,7`);
      }

      return Promise.reject('Invalid request');
    });
    //const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, period: 'D7',  output: 'json' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        //assert.equal(fileStub.called, true);
        //assert(cmdInstanceLogSpy.calledWith(`File saved to path '/Users/josephvelliah/Desktop/deviceusagedistributionusercounts.json'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });



  it('supports debug mode', () => {
    const options = mockCommand.options();
    let containsOption = false;
    options.forEach((o: any) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});