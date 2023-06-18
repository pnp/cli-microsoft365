import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./file-move');

describe(commands.FILE_MOVE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const stubAllPostRequests: any = (
    waitForJobResult: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/site/CreateCopyJobs') > -1) {
        return { value: [{ "EncryptionKey": "6G35dpTMegtzqT3rsZ/av6agpsqx/SUyaAHBs9fJE6A=", "JobId": "cee65dc5-8d05-41cc-8657-92a12d213f76", "JobQueueUri": "https://spobn1sn1m001pr.queue.core.windows.net:443/1246pq20180429-5305d83990eb483bb93e7356252715b4?sv=2014-02-14&sig=JUwFF1B0KVC2h0o5qieHPUG%2BQE%2BEhJHNpbzFf8QmCGc%3D&st=2018-04-28T07%3A00%3A00Z&se=2018-05-20T07%3A00%3A00Z&sp=rap" }] };
      }

      if ((opts.url as string).indexOf('/_api/site/GetCopyJobProgress') > -1) {
        if (waitForJobResult) {
          return waitForJobResult;
        }
        return {
          JobState: 0,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FilesCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Move\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFileStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFile\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        };
      }

      throw 'Invalid request';
    });
  };

  const stubAllGetRequests: any = (fileExists: any = null) => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('GetFileByServerRelativeUrl') > -1) {
        if (fileExists) {
          return fileExists;
        }
        return {};
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    (command as any).progressPollInterval = 0;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommand,
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_MOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['targetUrl', 'sourceUrl']);
  });

  it('should command complete successfully', async () => {
    stubAllPostRequests();
    stubAllGetRequests();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    });
    assert(loggerLogSpy.callCount === 0);
  });

  it('should complete successfully in 4 tries', async () => {
    let counter = 4;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        return;
      }

      if ((opts.url as string).indexOf('/_api/site/CreateCopyJobs') > -1) {
        return { value: [{ "EncryptionKey": "6G35dpTMegtzqT3rsZ/av6agpsqx/SUyaAHBs9fJE6A=", "JobId": "cee65dc5-8d05-41cc-8657-92a12d213f76", "JobQueueUri": "https://spobn1sn1m001pr.queue.core.windows.net:443/1246pq20180429-5305d83990eb483bb93e7356252715b4?sv=2014-02-14&sig=JUwFF1B0KVC2h0o5qieHPUG%2BQE%2BEhJHNpbzFf8QmCGc%3D&st=2018-04-28T07%3A00%3A00Z&se=2018-05-20T07%3A00%3A00Z&sp=rap" }] };
      }

      if ((opts.url as string).indexOf('/_api/site/GetCopyJobProgress') > -1) {
        // substract jobstate untill we hit jobstate = 0 (success)
        counter = (counter - 1);

        return {
          JobState: counter,
          Logs: ["{\r\n  \"Event\": \"JobEnd\",\r\n  \"JobId\": \"cee65dc5-8d05-41cc-8657-92a12d213f76\",\r\n  \"Time\": \"04/29/2018 22:00:08.370\",\r\n  \"FilesCreated\": \"1\",\r\n  \"BytesProcessed\": \"4860914\",\r\n  \"ObjectsProcessed\": \"2\",\r\n  \"TotalExpectedSPObjects\": \"2\",\r\n  \"TotalErrors\": \"0\",\r\n  \"TotalWarnings\": \"0\",\r\n  \"TotalRetryCount\": \"0\",\r\n  \"MigrationType\": \"Move\",\r\n  \"MigrationDirection\": \"Import\",\r\n  \"CreatedOrUpdatedFileStatsBySize\": \"{\\\"1-10M\\\":{\\\"Count\\\":1,\\\"TotalSize\\\":4860914,\\\"TotalDownloadTime\\\":24,\\\"TotalCreationTime\\\":2824}}\",\r\n  \"ObjectsStatsByType\": \"{\\\"SPUser\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":0,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPFile\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":3184,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0},\\\"SPListItem\\\":{\\\"Count\\\":1,\\\"TotalTime\\\":360,\\\"AccumulatedVersions\\\":0,\\\"ObjectsWithVersions\\\":0}}\",\r\n  \"TotalExpectedBytes\": \"4860914\",\r\n  \"CorrelationId\": \"8559629e-0036-5000-c38d-80b698e0cd79\"\r\n}"]
        };
      }

      throw 'Invalid request';
    });

    stubAllGetRequests();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc'
      }
    });
  });

  it('should succeed when run with option --deleteIfAlreadyExists and response 404', async () => {
    stubAllPostRequests();
    stubAllGetRequests();
    const fileDeleteError: any = {
      message: 'does not exist'
    };

    sinon.stub(Cli, 'executeCommand').rejects(fileDeleteError);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    });
  });

  it('should show error when recycleFile rejects with error', async () => {
    stubAllPostRequests();
    stubAllGetRequests();

    const fileDeleteError: any = {
      error: {
        message: 'Locked for use'
      },
      stderr: ''
    };

    sinon.stub(Cli, 'executeCommand').rejects(fileDeleteError);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    }), new CommandError(fileDeleteError.error.message));
  });

  it('should recycleFile format target url', async () => {
    const fileDeleteError: any = {
      error: {
        message: 'Invalid URL'
      },
      stderr: ''
    };

    sinon.stub(Cli, 'executeCommand').rejects(fileDeleteError);

    stubAllPostRequests();
    stubAllGetRequests();

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: '/abc/',
        deleteIfAlreadyExists: true
      }
    } as any), new CommandError(fileDeleteError.error.message));
  });

  it('should show error when getRequestDigestForSite rejects with error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    stubAllPostRequests();
    stubAllGetRequests();
    sinon.stub(Cli, 'executeCommand');
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').rejects(error);

    try {
      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          sourceUrl: 'abc/abc.pdf',
          targetUrl: 'abc',
          deleteIfAlreadyExists: true
        }
      } as any), new CommandError(error.error['odata.error'].message.value));
    }
    finally {
      sinonUtil.restore(spo.getRequestDigest);
      sinon.stub(spo, 'getRequestDigest').resolves({
        FormDigestValue: 'abc',
        FormDigestTimeoutSeconds: 1800,
        FormDigestExpiresAt: new Date(),
        WebFullUrl: 'https://contoso.sharepoint.com'
      });
    }
  });

  it('should show error when waitForJobResult rejects with JobError', async () => {
    const log = JSON.stringify({ Event: 'JobError', Message: 'error1' });

    stubAllPostRequests({ Logs: [log] });
    stubAllGetRequests();
    sinon.stub(Cli, 'executeCommand');
    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    } as any), new CommandError('error1'));
  });

  it('should show error when waitForJobResult rejects with JobFatalError', async () => {
    const log = JSON.stringify({ Event: 'JobFatalError', Message: 'error2' });

    stubAllPostRequests({ JobState: 0, Logs: [log] });
    stubAllGetRequests();

    sinon.stub(Cli, 'executeCommand');
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        sourceUrl: 'abc/abc.pdf',
        targetUrl: 'abc',
        deleteIfAlreadyExists: true
      }
    } as any), new CommandError('error2'));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: 'abc', targetUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetUrl: 'abc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
