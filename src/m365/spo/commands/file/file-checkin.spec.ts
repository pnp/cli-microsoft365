import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./file-checkin');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FILE_CHECKIN, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let stubPostResponses: any = (getFileByServerRelativeUrlResp: any = null, getFileByIdResp: any = null) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if (getFileByServerRelativeUrlResp) {
        return getFileByServerRelativeUrlResp;
      } else {
        if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativeUrl(') > -1) {
          return Promise.resolve();
        }
      }

      if (getFileByIdResp) {
        return getFileByIdResp;
      } else {
        if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      request.post,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_CHECKIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('command correctly handles file get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle checkin with url promise rejection',  (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File Not Found."}}});
    const getFileByServerRelativeUrlResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(getFileByServerRelativeUrlResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err.message), JSON.stringify(expectedError));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle checkin with id promise rejection',  (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File Not Found."}}});
    const getFileByIdResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, getFileByIdResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err.message), JSON.stringify(expectedError));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct API url when UniqueId option is passed', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment=\'\',checkintype=1)');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call "DONE" when in verbose', (done) => {
    stubPostResponses();

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0], 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct API url when URL option is passed', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        fileUrl: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=1)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct API url when tenant root URL option is passed', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        fileUrl: '/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=1)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call correctly the API when type is minor', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        fileUrl: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'minor'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=0)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call correctly the API when type is overwrite', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        fileUrl: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'overwrite'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=2)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call correctly the API when comment specified', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        fileUrl: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        comment: 'abc1'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='abc1',checkintype=1)");
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('should call correctly the API when type is minor (id)', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'minor'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='',checkintype=0)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call correctly the API when type is overwrite (id)', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'overwrite'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='',checkintype=2)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call correctly the API when comment specified (id)', (done) => {
    const postStub: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: false,
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        comment: 'abc1'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='abc1',checkintype=1)");
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id or url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when comment lenght more than 1023', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', comment: 'ayMfuJBMDa3y3y8qitRb4U6VSbBVjeIxno45Ws6baZ1uatvxGVDS98zQu88QUjyeYXDbLey1dWTKdgMGw4LGeqfn080VszB5vMCrBEAnLYT54E94eW3YQe67Ub92oD0DG0U8gxMQJ0SWdVG9m5R5dL31YWx1Y5OH8KMtoAFkfo2lnbHVBMnCiO8oyuiRzVbTLkZB7mdih3F74ck3kEM7Lr1ayXkwHKK5h9MnTcVTWZVXafMOsuLYaVnUB7auhaamQ4JMBUFNpKhCjrNQVlYz0NlwJimlk5tPeR6crgeCm3u4YJtc1dBL2Ex7FRfvJ4g44WnkPLyU3PIXrHTjZtlgOKn4m9BiABuwznqiuytCcKbLxaTQcqHsbC7w20vnZxnLHYNnqXeDqwf6o43Si6duSeIZSixwoK4nE8qpCZk36jkwZBXASuv5aOyWLOsD19JjK7Ev3567b6oo11krIOpd0TSRihphELWnk9A71xpkCN1ljmSTnrITgQ7AxIaWOHvBIv5Swffi6AUM2DeLyz61EVe0fgAdVU3UySGSHGmUJEGqVBUlX7zZw2xSWswgvQphziHp2sKcnONWaaeDvbr27g67HrkkyYO3z0R5nY9TdSfkqDDQVSFdM3Sd6WLRKKKn64pcUzo9NcFNKzMSvRR0FbZFirpEcIfTCrSLaIiRZYCoGdj0BfePz83cimDmlVWS87UXugXmeWNpKTqQ1qG9y0fMwGIxFory4YbeRP9vKqX0vueCGKErb7tItC09jpLp0J8yMaj0iDdZ83Yc3JHunVmqZh56hmUroU8ER6ApPS3oDooEGH59e5I4DU8LG4rpAPmECX6oC8w9eZfM7U0uugT9Yx2ZAoDwvk0jYJz8SuU6dL6aFYtf7wzYcBcjc8gBySbeVZYPoLE3TGP1A0K8HNiZavHjsJWK0GIYVDT4QEsJO4R9PykRkn0O6TyDkaIgqju9hV7lqy9YqKawvBAUlNyK7b01fkra5UBrZzYz83k3OYWmG2naAcKuNuPs7OJ6' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation wrong type', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', type: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', fileUrl: '/sites/project-x/documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation id, type and comment params correct', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', type: 'overwrite', comment: '123' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation url, type and comment params correct', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', fileUrl: '/sites/docs/abc.txt', type: 'overwrite', comment: '123' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation type is major', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', fileUrl: '/sites/docs/abc.txt', type: 'major' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation type is minor', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', fileUrl: '/sites/docs/abc.txt', type: 'minor' } });
    assert.strictEqual(actual, true);
  });
});
