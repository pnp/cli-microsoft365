import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./folder-move');

describe(commands.FOLDER_MOVE, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

  let stubAllPostRequests: any = (
    recycleFolder: any = null,
    createCopyJobs: any = null,
    getCopyJobProgress: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        if (recycleFolder) {
          return recycleFolder;
        }
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('/_api/site/CreateCopyJobs') > -1) {
        if (createCopyJobs) {
          return createCopyJobs;
        }
        return Promise.resolve({ value: [{ "EncryptionKey": "6G35dpTMegtzqT3rsZ/av6agpsqx/SUyaAHBs9fJE6A=", "JobId": "cee65dc5-8d05-41cc-8657-92a12d213f76", "JobQueueUri": "https://spobn1sn1m001pr.queue.core.windows.net:443/1246pq20180429-5305d83990eb483bb93e7356252715b4?sv=2014-02-14&sig=JUwFF1B0KVC2h0o5qieHPUG%2BQE%2BEhJHNpbzFf8QmCGc%3D&st=2018-04-28T07%3A00%3A00Z&se=2018-05-20T07%3A00%3A00Z&sp=rap" }] });
      }

      if ((opts.url as string).indexOf('/_api/site/GetCopyJobProgress') > -1) {
        if (getCopyJobProgress) {
          return getCopyJobProgress;
        }
        return Promise.resolve({
          JobState: 0,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Move\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  let stubAllGetRequests: any = (folderExists: any = null) => {

    return sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('GetFolderByServerRelativeUrl') > -1) {
        if (folderExists) {
          return folderExists;
        }
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
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
      appInsights.trackEvent,
      global.setTimeout
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_MOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should command complete successfully (verbose)', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(loggerSpy.lastCall.args[0] === 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should command complete successfully', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(loggerSpy.callCount === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should complete successfully in 4 tries. ', (done) => {
    var counter = 4;
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('/_api/site/CreateCopyJobs') > -1) {
        return Promise.resolve({ value: [{ "EncryptionKey": "6G35dpTMegtzqT3rsZ/av6agpsqx/SUyaAHBs9fJE6A=", "JobId": "cee65dc5-8d05-41cc-8657-92a12d213f76", "JobQueueUri": "https://spobn1sn1m001pr.queue.core.windows.net:443/1246pq20180429-5305d83990eb483bb93e7356252715b4?sv=2014-02-14&sig=JUwFF1B0KVC2h0o5qieHPUG%2BQE%2BEhJHNpbzFf8QmCGc%3D&st=2018-04-28T07%3A00%3A00Z&se=2018-05-20T07%3A00%3A00Z&sp=rap" }] });
      }

      if ((opts.url as string).indexOf('/_api/site/GetCopyJobProgress') > -1) {
        // substract jobstate untill we hit jobstate = 0 (success)
        counter = (counter - 1);

        return Promise.resolve({
          JobState: counter,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Move\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        });
      }

      return Promise.reject('Invalid request');

    });

    stubAllGetRequests();

    command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(loggerSpy.lastCall.args[0] === 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error when getCopyJobProgress rejects with JobError', (done) => {
    const getCopyJobProgress = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobError', Message: 'error1' });
      return resolve({ Logs: [log] });
    });
    stubAllPostRequests(null, null, getCopyJobProgress);
    stubAllGetRequests();

    command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error when getCopyJobProgress rejects with JobFatalError', (done) => {
    const getCopyJobProgress = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobFatalError', Message: 'error2' });
      return resolve({ JobState: 0, Logs: [log] });
    });
    stubAllPostRequests(null, null, getCopyJobProgress);
    stubAllGetRequests();

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error2')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should complete successfully where baseUrl has a trailing /', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      exportObjectUris: [
        'https://contoso.sharepoint.com/sites/team-a/library/folder1'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true,
        'IsMoveMode': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/folder1',
        targetUrl: 'sites/team-b/library2'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should complete successfully where sourceUrl and targetUrl has a trailing /', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      exportObjectUris: [
        'https://contoso.sharepoint.com/sites/team-a/library/folder1'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true,
        'IsMoveMode': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/folder1/',
        targetUrl: 'sites/team-b/library2/'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should complete successfully where sourceUrl and targetUrl has a beginning /', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      exportObjectUris: [
        'https://contoso.sharepoint.com/sites/team-a/library/folder1'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true,
        'IsMoveMode': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: '/library/folder1/',
        targetUrl: '/sites/team-b/library2/'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.strictEqual(actual, true);
  });
});
