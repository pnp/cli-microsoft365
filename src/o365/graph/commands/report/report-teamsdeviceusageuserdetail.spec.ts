import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./report-teamsdeviceusageuserdetail');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import request from '../../../../request';

describe(commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
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
    assert.equal(command.name.startsWith(commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if both period and date options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if both period and date options set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { period: 'D7', date: '2019-05-01' } });
    assert.notEqual(actual, true);
  });

  it('fails validation on invalid period', () => {
    const actual = (command.validate() as CommandValidate)({ options: { period: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation on valid \'D7\' period', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        period: 'D7'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'D30\' period', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        period: 'D30'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'D90\' period', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        period: 'D90'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'180\' period', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        period: 'D90'
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if the date option is not a valid date string', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        date: '2018-X-09'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the date option is a valid date', () => {
    const actual = (command.validate() as CommandValidate)({ options: { date: '2019-05-01' } });
    assert(actual);
  });

  it('gets details about Microsoft Teams device usage by user for the given period', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D7')`) {
        return Promise.resolve('Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Used Web,Used Windows Phone,Used iOS,Used Mac,Used Android Phone,Used Windows,Report Period');
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, () => {
      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Microsoft Teams device usage by user for the given date', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(date='2019-05-01')`) {
        return Promise.resolve({
          "value": 'Report Refresh Date,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Used Web,Used Windows Phone,Used iOS,Used Mac,Used Android Phone,Used Windows,Report Period'
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, date: '2019-05-01' } }, () => {
      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false, date: '2019-05-01' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.REPORT_TEAMSDEVICEUSAGEUSERDETAIL));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});