import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./file-rename');

describe(commands.FILE_RENAME, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  let stubAllPostRequests: any = (
    recycleFile: any = null,
    MoveTo: any = null,
   
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        if (recycleFile) {
          return recycleFile;
        }
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('/MoveTo()') > -1) {
        if (MoveTo) {
          return MoveTo;
        }
        return Promise.resolve();
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
    assert.strictEqual(command.name.startsWith(commands.FILE_RENAME), true);
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
        targetFileName: 'abc.pdf'
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.lastCall.args[0] === 'DONE');
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
        targetFileName: 'abc.pdf'
      }
    }, () => {
      try {
        assert(loggerLogSpy.callCount === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should complete MoveTo. ', (done) => {
    
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('/MoveTo()') > -1) {
        return Promise.resolve();
      }

     
      return Promise.reject('Invalid request');
    });

    stubAllGetRequests();

    command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf'
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.lastCall.args[0] === 'DONE');
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

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File not found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed when run with option --force', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.lastCall.calledWith('DONE'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed when run with option --force and response 404', (done) => {
    const recycleFile404 = new Promise<any>((resolve, reject) => {
      return reject({ statusCode: 404 });
    });
    stubAllPostRequests(recycleFile404);
    stubAllGetRequests();

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
      }
    }, () => {
      try {
        assert(loggerLogToStderrSpy.lastCall.calledWith('DONE'));
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

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
      }
    } as any, (err?: any) => {
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

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
      }
    } as any, (err?: any) => {
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

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
      }
    } as any, (err?: any) => {
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

    command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
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

  it('should show error when waitForJobResult rejects with JobFatalError', (done) => {
    const waitForJobResult = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobFatalError', Message: 'error2' });
      return resolve({ JobState: 0, Logs: [log] });
    });
    stubAllPostRequests(null, null, waitForJobResult);
    stubAllGetRequests();

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetFileName: 'abc.pdf',
        force: true
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
        'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf'
      ],
      destinationUri: 'https://contoso.sharepoint.com/sites/team-b/library2',
      options: {
        'AllowSchemaMismatch': false,
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });
    stubAllGetRequests();

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/file1.pdf',
        targetFuileName: 'file2.pdf'
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
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    stubAllGetRequests();

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/file1.pdf/',
        targetFileName: 'file2-pdf'
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
      actual = JSON.stringify(opts.data);
      if (
        opts.data.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/file1.pdf' &&
        opts.data.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    stubAllGetRequests();

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: '/library/file1.pdf/',
        targetFileName: 'file2.pdf'
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
    const actual = command.validate({ options: { webUrl: 'foo', sourceUrl: 'abc', targetFileName: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetFileName: 'abc' } });
    assert.strictEqual(actual, true);
  });
});
