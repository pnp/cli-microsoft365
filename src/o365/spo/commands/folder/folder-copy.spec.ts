import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./folder-copy');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.FOLDER_COPY, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  let stubAllPostRequests: any = (
    getAccessToken = null,
    getRequestDigestForSite = null,
    recycleFolder = null,
    createCopyJobs = null,
    getCopyJobProgress = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {

      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        if (getAccessToken) {
          return getAccessToken;
        }
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (getRequestDigestForSite) {
          return getRequestDigestForSite;
        }
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      if (opts.url.indexOf('/recycle()') > -1) {
        if (recycleFolder) {
          return recycleFolder;
        }
        return Promise.resolve();
      }

      if (opts.url.indexOf('/_api/site/CreateCopyJobs') > -1) {
        if (createCopyJobs) {
          return createCopyJobs;
        }
        return Promise.resolve({ value: [{ "EncryptionKey": "6G35dpTMegtzqT3rsZ/av6agpsqx/SUyaAHBs9fJE6A=", "JobId": "cee65dc5-8d05-41cc-8657-92a12d213f76", "JobQueueUri": "https://spobn1sn1m001pr.queue.core.windows.net:443/1246pq20180429-5305d83990eb483bb93e7356252715b4?sv=2014-02-14&sig=JUwFF1B0KVC2h0o5qieHPUG%2BQE%2BEhJHNpbzFf8QmCGc%3D&st=2018-04-28T07%3A00%3A00Z&se=2018-05-20T07%3A00%3A00Z&sp=rap" }] });
      }

      if (opts.url.indexOf('/_api/site/GetCopyJobProgress') > -1) {
        if (getCopyJobProgress) {
          return getCopyJobProgress;
        }
        return Promise.resolve({
          JobState: 0,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FoldersCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Copy\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFolderStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFolder\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  let stubAllGetRequests: any = (folderExists = null) => {

    return sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('GetFolderByServerRelativeUrl') > -1) {
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
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
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
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FOLDER_COPY), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.FOLDER_COPY);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should command complete successfully (verbose)', (done) => {
    stubAllPostRequests();
    stubAllGetRequests();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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

  it('should show error when getCopyJobProgress rejects with JobError', (done) => {
    const getCopyJobProgress = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobError', Message: 'error1' });
      return resolve({ Logs: [log] });
    });
    stubAllPostRequests(null, null, null, null, getCopyJobProgress);
    stubAllGetRequests();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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

  it('should show error when getCopyJobProgress rejects with JobFatalError', (done) => {
    const getCopyJobProgress = new Promise<any>((resolve, reject) => {
      const log = JSON.stringify({ Event: 'JobFatalError', Message: 'error2' });
      return resolve({ JobState: 0, Logs: [log] });
    });
    stubAllPostRequests(null, null, null, null, getCopyJobProgress);
    stubAllGetRequests();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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

  it('should setTimeout when getCopyJobProgress JobState is not 0', (done) => {
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
      (command as any).getCopyJobProgress(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert(resp === undefined);
        postRequests.restore();
        done();
      }, (e: any) => {
        assert.fail('getCopyJobProgress couldn\'t resolve.');
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
  });

  it('should setTimeout when getCopyJobProgress reject, but retry limit not reached', (done) => {
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
      (command as any).getCopyJobProgress(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert(resp === undefined);
        postRequests.restore();
        done();
      }, (e: any) => {
        assert.fail('getCopyJobProgress couldn\'t resolve.');
        postRequests.restore();
        done();
      });
    }
    catch (e) {
      done(e);
    }
  });

  it('should show error when getCopyJobProgress reject and retry limit reached', (done) => {
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
      (command as any).getCopyJobProgress(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert.fail('getCopyJobProgress shouldn\'t resolve, but reject.');
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

  it('should getCopyJobProgress timeout', (done) => {
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
      (command as any).getCopyJobProgress(jobProgressOptions, cmdInstance).then((resp: any) => {
        assert.fail('getCopyJobProgress shouldn\'t resolve, but reject.');
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

  it('should combine url with baseUrl that last char is /', () => {
    const actual = (command as any).urlCombine('https://contoso.com/', 'sites/abc');
    assert.equal(actual, 'https://contoso.com/sites/abc');
  });

  it('should combine url with relativeUrl that last char is /', () => {
    const actual = (command as any).urlCombine('https://contoso.com', 'sites/abc/');
    assert.equal(actual, 'https://contoso.com/sites/abc');
  });

  it('should combine url with relativeUrl that first char is /', () => {
    const actual = (command as any).urlCombine('https://contoso.com/', '/sites/abc/');
    assert.equal(actual, 'https://contoso.com/sites/abc');
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});