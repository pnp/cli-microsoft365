import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./file-copy');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FILE_COPY, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let stubAllPostRequests: any = (
    recycleFile: any = null,
    createCopyJobs: any = null,
    waitForJobResult: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        if (recycleFile) {
          return recycleFile;
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
        if (waitForJobResult) {
          return waitForJobResult;
        }
        return Promise.resolve({
          JobState: 0,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FilesCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFileStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFile\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  let stubAllGetRequests: any = (fileExists: any = null) => {
    return sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('GetFileByServerRelativeUrl') > -1) {
        if (fileExists) {
          return fileExists;
        }
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub((command as any), 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent,
      global.setTimeout
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_COPY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should command complete successfully (verbose)', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.args[0] === 'DONE');
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

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.callCount === 0);
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
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FilesCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Move\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFileStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFile\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        });
      }

      return Promise.reject('Invalid request');
    });

    stubAllGetRequests();

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.args[0] === 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should fail if source file not found', (done) => {
    stubAllPostRequests();
    const rejectFileExists = new Promise<any>((resolve, reject) => {
      return reject('File not found.');
    });
    stubAllGetRequests(rejectFileExists);

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File not found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed when run with option --deleteIfAlreadyExists', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith('DONE'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed when run with option --deleteIfAlreadyExists and response 404', (done) => {
    const recycleFile404 = new Promise<any>((resolve, reject) => {
      return reject({ statusCode: 404 });
    });
    stubAllPostRequests(recycleFile404);
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith('DONE'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error when recycleFile rejects with error', (done) => {
    const recycleFile = new Promise<any>((resolve, reject) => {
      return reject('abc');
    });
    stubAllPostRequests(recycleFile);
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should recycleFile format target url', (done) => {
    const recycleFile = new Promise<any>((resolve, reject) => {
      return reject('abc');
    });
    stubAllPostRequests(recycleFile);
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: '/abc/',
        deleteIfAlreadyExists: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('abc')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error when getRequestDigest rejects with error', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();
    Utils.restore((command as any).getRequestDigest);
    sinon.stub((command as any), 'getRequestDigest').callsFake(() => Promise.reject('error'));

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore((command as any).getRequestDigest);
        sinon.stub((command as any), 'getRequestDigest').callsFake(() => Promise.resolve({
          FormDigestValue: 'abc'
        }));
      }
    });
  });

  it('should show error when waitForJobResult rejects with JobError', (done) => {
    const waitForJobResult = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobError', Message: 'error1' });
      return resolve({ Logs: [log] });
    });
    stubAllPostRequests(null, null, waitForJobResult);
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error when waitForJobResult rejects with JobFatalError', (done) => {
    const waitForJobResult = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobFatalError', Message: 'error2' });
      return resolve({ JobState: 0, Logs: [log] });
    });
    stubAllPostRequests(null, null, waitForJobResult);
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }, (err?: any) => {
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
        'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });
    stubAllGetRequests();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/file1.pdf',
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
        'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    stubAllGetRequests();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/file1.pdf/',
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
        'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    stubAllGetRequests();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: '/library/file1.pdf/',
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.strictEqual(actual, true);
  });
});
