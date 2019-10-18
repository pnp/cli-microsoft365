import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-report-activitydetail');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import request from '../../../../request';
import * as fs from 'fs';

describe(commands.O365GROUP_REPORT_ACTIVITYDETAIL, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let writeFileSyncFake = () => { };

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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
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
    assert.equal(command.name.startsWith(commands.O365GROUP_REPORT_ACTIVITYDETAIL), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if both period and date options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if both period and date options set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { period: 'D7', date: '2019-09-28' } });
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
        period: 'D180'
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
    const actual = (command.validate() as CommandValidate)({ options: { date: '2019-09-28' } });
    assert(actual);
  });

  it('fails validation if specified outputFile directory path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({
      options: {
        period: 'D7',
        outputFile: '/path/not/found.zip'
      }
    });
    Utils.restore(fs.existsSync);
    assert.notEqual(actual, true);
  });

  it('gets details about Office 365 Groups activity by group for the given period', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, period: 'D7' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group for the given date', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(date=2019-09-28)`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, date: '2019-09-28' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(date=2019-09-28)");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group for the given period and export report data in txt format', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7', outputFile: './o365groupactivitydetail.txt' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group when output is json', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });

    const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7', output: 'json' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        assert.equal(fileStub.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group for the given period and export report data in txt format with output', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });
    const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7', outputFile: './o365groupactivitydetail.txt', output: 'text' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group for the given period and export report data in json format', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve(`
        Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period
        2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7
        2019-10-01,D.Marketing,True,user2@sharepointrider.onmicrosoft.com,2019-05-30,Private,4,0,,,,,,413,3882649,4,1596856,02826124-adbe-4d57-8ccb-a2b5647cad14,7
        `);
      }

      return Promise.reject('Invalid request');
    });
    const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, period: 'D7', outputFile: './o365groupactivitydetail.json' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        assert.equal(fileStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets details about Office 365 Groups activity by group for the given period and export report data in json format with output', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')`) {
        return Promise.resolve('Report Refresh Date,Group Display Name,Is Deleted,Owner Principal Name,Last Activity Date,Group Type,Member Count,External Member Count,Exchange Received Email Count,SharePoint Active File Count,Yammer Posted Message Count,Yammer Read Message Count,Yammer Liked Message Count,Exchange Mailbox Total Item Count,Exchange Mailbox Storage Used (Byte),SharePoint Total File Count,SharePoint Site Storage Used (Byte),Group Id,Report Period\n2019-10-01,Pavithra Library,False,user1@sharepointrider.onmicrosoft.com,,Private,7,2,,,,,,430,4757931,0,1450329,01c48e08-ff4a-4d47-bb42-947581d1b3fe,7');
      }

      return Promise.reject('Invalid request');
    });
    const fileStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, period: 'D7', outputFile: './o365groupactivitydetail.json', output: 'json' } }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D7')");
        assert.equal(requestStub.lastCall.args[0].headers["accept"], 'application/json;odata.metadata=none');
        assert.equal(requestStub.lastCall.args[0].json, true);
        assert.equal(fileStub.called, true);
        assert(cmdInstanceLogSpy.calledWith(`File saved to path './o365groupactivitydetail.json'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false, date: '2019-09-28' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying outputFile', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--outputFile') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
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
    assert(find.calledWith(commands.O365GROUP_REPORT_ACTIVITYDETAIL));
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