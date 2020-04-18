import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./folder-copy');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FOLDER_COPY, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let stubAllPostRequests: any = (
    recycleFolder: any = null,
    createCopyJobs: any = null,
    waitForJobResult: any = null
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
        if (waitForJobResult) {
          return waitForJobResult;
        }
        return Promise.resolve({
          JobState: 0,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
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
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FOLDER_COPY), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        targetUrl: 'abc'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
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
        targetUrl: 'abc'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('error2')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should setTimeout when waitForJobResult JobState is not 0', (done) => {
    const postRequests = sinon.stub(request, 'post');
    postRequests.onFirstCall().resolves({
      JobState: 4,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:29.945\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Export\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{}\",\r\n  \"ObjectsStatsByType\": \"{}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}", "{\r\n  \"Event\": \"JobLogFolderCreate\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:30.539\",\r\n  \"FolderName\": \"Import-7be2c14c-b998-4b30-9b43-c2be0f95d8b9-0.log\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}", "{\r\n  \"Event\": \"JobStart\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:30.570\",\r\n  \"SiteId\": \"956c8970-f858-42ac-a06f-bbdca4a0374b\",\r\n  \"WebId\": \"d6d96969-217f-4306-b15b-fe35b6b754cc\",\r\n  \"DBId\": \"eb30ff26-a12c-431e-bb10-68fdac21ce28\",\r\n  \"FarmId\": \"67b76b49-9245-4dfc-a1f7-b4503cf6ea69\",\r\n  \"ServerId\": \"2a00b725-2871-4e42-98fd-e41c577ed494\",\r\n  \"SubscriptionId\": \"ea1787c6-7ce2-4e71-be47-5e0deb30f9e4\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}"]
    });

    postRequests.onSecondCall().resolves({
      JobState: 0,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
    });

    const jobProgressOptions: any = {
      webUrl: 'https://contoso.sharepoint.com',
      accessToken: 'abc',
      copyJopInfo: 'abc',
      progressMaxPollAttempts: 2,
      progressPollInterval: 0,
      progressRetryAttempts: 5
    }
    const log: any = [];
    const cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };

    try {
      (command as any).waitForJobResult(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert(resp === undefined);
        postRequests.restore();
        done();
      }, (e: any) => {
        assert.fail('waitForJobResult couldn\'t resolve.');
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
  });

  it('should setTimeout when waitForJobResult reject, but retry limit not reached', (done) => {
    const postRequests = sinon.stub(request, 'post');
    // GetCopyJobProgress reject
    postRequests.onFirstCall().rejects('error');
    // GetCopyJobProgress #2 JobState = 0
    postRequests.onSecondCall().resolves({
      JobState: 0,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
    });

    const jobProgressOptions: any = {
      webUrl: 'https://contoso.sharepoint.com',
      accessToken: 'abc',
      copyJopInfo: 'abc',
      progressMaxPollAttempts: 2,
      progressPollInterval: 0,
      progressRetryAttempts: 1
    }
    const log: any = [];
    const cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };

    try {
      (command as any).waitForJobResult(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert(resp === undefined);
        postRequests.restore();
        done();
      }, (e: any) => {
        assert.fail('waitForJobResult couldn\'t resolve.');
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
  });

  it('should show error when waitForJobResult reject and retry limit reached', (done) => {
    const postRequests = sinon.stub(request, 'post');
    postRequests.onFirstCall().rejects('error');
    postRequests.onSecondCall().rejects('error');
    postRequests.onThirdCall().resolves({
      JobState: 0,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
    });

    const jobProgressOptions: any = {
      webUrl: 'https://contoso.sharepoint.com',
      accessToken: 'abc',
      copyJopInfo: 'abc',
      progressMaxPollAttempts: 2,
      progressPollInterval: 0,
      progressRetryAttempts: 1
    }
    const log: any = [];
    const cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };

    try {
      (command as any).waitForJobResult(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert.fail('waitForJobResult shouldn\'t resolve, but reject.');
        postRequests.restore();
        done();
      }, (e: any) => {
        assert(e !== undefined);
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
  });

  it('should waitForJobResult timeout', (done) => {
    const postRequests = sinon.stub(request, 'post');
    // GetCopyJobProgress #1 JobState = 4
    postRequests.onFirstCall().resolves({
      JobState: 4,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:29.945\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Export\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{}\",\r\n  \"ObjectsStatsByType\": \"{}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}", "{\r\n  \"Event\": \"JobLogFolderCreate\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:30.539\",\r\n  \"FolderName\": \"Import-7be2c14c-b998-4b30-9b43-c2be0f95d8b9-0.log\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}", "{\r\n  \"Event\": \"JobStart\",\r\n  \"JobId\": \"7be2c14c-b998-4b30-9b43-c2be0f95d8b9\",\r\n  \"Time\": \"04/29/2018 23:39:30.570\",\r\n  \"SiteId\": \"956c8970-f858-42ac-a06f-bbdca4a0374b\",\r\n  \"WebId\": \"d6d96969-217f-4306-b15b-fe35b6b754cc\",\r\n  \"DBId\": \"eb30ff26-a12c-431e-bb10-68fdac21ce28\",\r\n  \"FarmId\": \"67b76b49-9245-4dfc-a1f7-b4503cf6ea69\",\r\n  \"ServerId\": \"2a00b725-2871-4e42-98fd-e41c577ed494\",\r\n  \"SubscriptionId\": \"ea1787c6-7ce2-4e71-be47-5e0deb30f9e4\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CorrelationId\": \"355f629e-707d-5000-634c-4c5cdd1e62d2\"\r\n}"]
    });
    // GetCopyJobProgress #2 JobState = 0
    postRequests.onSecondCall().resolves({
      JobState: 4,
      Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
    });

    const jobProgressOptions: any = {
      webUrl: 'https://contoso.sharepoint.com',
      accessToken: 'abc',
      copyJopInfo: 'abc',
      progressMaxPollAttempts: 1,
      progressPollInterval: 0,
      progressRetryAttempts: 5
    }
    const log: any = [];
    const cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };

    try {
      (command as any).waitForJobResult(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert.fail('waitForJobResult shouldn\'t resolve, but reject.');
        postRequests.restore();
        done();
      }, (e: any) => {
        assert(e !== undefined);
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
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
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/folder1',
        targetUrl: 'sites/team-b/library2'
      }
    }, () => {
      try {
        assert.equal(actual, expected);
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
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: 'library/folder1/',
        targetUrl: 'sites/team-b/library2/'
      }
    }, () => {
      try {
        assert.equal(actual, expected);
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
        'IgnoreVersionHistory': true
      }
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      actual = JSON.stringify(opts.body);
      if (
        opts.body.exportObjectUris[0] === 'https://contoso.sharepoint.com/sites/team-a/library/folder1' &&
        opts.body.destinationUri === 'https://contoso.sharepoint.com/sites/team-b/library2' &&
        opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/site/CreateCopyJobs'
      ) {
        return Promise.resolve();

      }
      return Promise.reject('Invalid request');

    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team-a/',
        sourceUrl: '/library/folder1/',
        targetUrl: '/sites/team-b/library2/'
      }
    }, () => {
      try {
        assert.equal(actual, expected);
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

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.equal(actual, true);
  });

  it('fails validation if the sourceUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the targetUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc' } });
    assert.notEqual(actual, true);
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
    assert(find.calledWith(commands.FOLDER_COPY));
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