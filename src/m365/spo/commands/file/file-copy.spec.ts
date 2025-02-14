import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './file-copy.js';
import { settingsNames } from '../../../../settingsNames.js';
import { CreateFileCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';

describe(commands.FILE_COPY, () => {
  const sourceWebUrl = 'https://contoso.sharepoint.com/sites/Sales';
  const sourceDocumentName = 'Document.pdf';
  const sourceServerRelUrl = '/sites/Sales/Shared Documents/' + sourceDocumentName;
  const sourceAbsoluteUrl = 'https://contoso.sharepoint.com' + sourceServerRelUrl;
  const sourceDocId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';

  const destWebUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const destSiteRelUrl = '/Documents/Logos';
  const destServerRelUrl = '/sites/Marketing' + destSiteRelUrl;
  const destAbsoluteTargetUrl = 'https://contoso.sharepoint.com' + destServerRelUrl;
  const destDocId = '15488d89-b82b-40be-958a-922b2ed79383';

  const copyJobInfo = {
    EncryptionKey: '2by8+2oizihYOFqk02Tlokj8lWUShePAEE+WMuA9lzA=',
    JobId: 'd812e5a0-d95a-4e4f-bcb7-d4415e88c8ee',
    JobQueueUri: 'https://spoam1db1m020p4.queue.core.windows.net/2-1499-20240831-29533e6c72c6464780b756c71ea3fe92?sv=2018-03-28&sig=aX%2BNOkUimZ3f%2B%2BvdXI95%2FKJI1e5UE6TU703Dw3Eb5c8%3D&st=2024-08-09T00%3A00%3A00Z&se=2024-08-31T00%3A00%3A00Z&sp=rap',
    SourceListItemUniqueIds: [
      sourceDocId
    ]
  };

  const copyJobResult = {
    Event: 'JobFinishedObjectInfo',
    JobId: '6d1eda82-0d1c-41eb-ab05-1d9cd4afe786',
    Time: '08/10/2024 18:59:40.145',
    SourceObjectFullUrl: sourceAbsoluteUrl,
    TargetServerUrl: 'https://contoso.sharepoint.com',
    TargetSiteId: '794dada8-4389-45ce-9559-0de74bf3554a',
    TargetWebId: '8de9b4d3-3c30-4fd0-a9d7-2452bd065555',
    TargetListId: '44b336a5-e397-4e22-a270-c39e9069b123',
    TargetObjectUniqueId: destDocId,
    TargetObjectSiteRelativeUrl: destSiteRelUrl.substring(1),
    CorrelationId: '5efd44a1-c034-9000-9692-4e1a1b3ca33b'
  };

  const destFileResponse = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{C194762B-3F54-4F5F-9F5C-EBA26084E29D},53,23',
    CustomizedPageStatus: 0,
    ETag: '"{C194762B-3F54-4F5F-9F5C-EBA26084E29D},53"',
    Exists: true,
    ExistsAllowThrowForPolicyFailures: true,
    ExistsWithException: true,
    IrmEnabled: false,
    Length: '18911',
    Level: 1,
    LinkingUri: `${destAbsoluteTargetUrl + '/' + sourceDocumentName}?d=wc194762b3f544f5f9f5ceba26084e29d`,
    LinkingUrl: '',
    MajorVersion: 14,
    MinorVersion: 0,
    Name: sourceDocumentName,
    ServerRelativeUrl: destServerRelUrl + '/' + sourceDocumentName,
    TimeCreated: '2024-05-01T19:54:50Z',
    TimeLastModified: '2024-08-10T19:31:34Z',
    Title: '',
    UIVersion: 7168,
    UIVersionLabel: '14.0',
    UniqueId: destDocId
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;

  let spoUtilCreateCopyJobStub: sinon.SinonStub;
  let spoUtilGetCopyJobResultStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => settingName === settingsNames.prompt ? false : defaultValue);
    spoUtilCreateCopyJobStub = sinon.stub(spo, 'createFileCopyJob').resolves(copyJobInfo);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };

    loggerLogSpy = sinon.spy(logger, 'log');
    spoUtilGetCopyJobResultStub = sinon.stub(spo, 'getCopyJobResult').resolves(copyJobResult);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      spo.getCopyJobResult
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: sourceAbsoluteUrl, targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the sourceId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceId: 'invalid', targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the sourceId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceId: sourceDocId, targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both sourceId and sourceUrl options are not specified', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both sourceId and url options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceId: sourceDocId, sourceUrl: sourceAbsoluteUrl, targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if nameConflictBehavior has an invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceUrl: sourceAbsoluteUrl, targetUrl: destAbsoluteTargetUrl, nameConflictBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are provided', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceUrl: sourceAbsoluteUrl, targetUrl: destAbsoluteTargetUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required options are provided with optional', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceUrl: sourceAbsoluteUrl, targetUrl: destAbsoluteTargetUrl, ignoreVersionHistory: true, nameConflictBehavior: 'RePlAcE' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly outputs exactly one result when file is copied when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFileById('${sourceDocId}')/ServerRelativePath`) {
        return {
          DecodedUrl: destAbsoluteTargetUrl + `/${sourceDocumentName}`
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceId: sourceDocId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert(loggerLogSpy.calledOnce);
  });

  it('correctly outputs result when file is copied when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFileById('${sourceDocId}')/ServerRelativePath`) {
        return {
          DecodedUrl: sourceAbsoluteUrl
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceId: sourceDocId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], destFileResponse);
  });

  it('correctly outputs exactly one result when file is copied when using sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert(loggerLogSpy.calledOnce);
  });

  it('correctly outputs result when file is copied when using sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], destFileResponse);
  });

  it('correctly copies a file when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFileById('${sourceDocId}')/ServerRelativePath`) {
        return {
          DecodedUrl: sourceAbsoluteUrl
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceId: sourceDocId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Fail,
        bypassSharedLock: false,
        ignoreVersionHistory: false,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a file when using sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'fail'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Fail,
        bypassSharedLock: false,
        ignoreVersionHistory: false,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a file when using sourceUrl with extra options', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'rename',
        bypassSharedLock: true,
        ignoreVersionHistory: true,
        newName: 'Document-renamed.pdf'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Rename,
        bypassSharedLock: true,
        ignoreVersionHistory: true,
        operation: 'copy',
        newName: 'Document-renamed.pdf'
      }
    ]);
  });

  it('correctly copies a file when using sourceUrl with new name without extension', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'replace',
        newName: 'Document-renamed'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Replace,
        bypassSharedLock: false,
        ignoreVersionHistory: false,
        operation: 'copy',
        newName: 'Document-renamed.pdf'
      }
    ]);
  });

  it('correctly polls for the copy job to finish', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFileById('${destDocId}')`) {
        return destFileResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert.deepStrictEqual(spoUtilGetCopyJobResultStub.lastCall.args, [
      sourceWebUrl,
      copyJobInfo
    ]);
  });

  it('outputs no result when skipWait is specified', async () => {
    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        skipWait: true
      }
    });

    assert(loggerLogSpy.notCalled);
  });

  it('correctly skips polling when skipWait is specified', async () => {
    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        skipWait: true
      }
    });

    assert(spoUtilGetCopyJobResultStub.notCalled);
  });

  it('correctly handles error when sourceId does not exist', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-2147024894, System.IO.FileNotFoundException',
          message: {
            lang: 'en-US',
            value: 'File Not Found.'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceId: sourceDocId,
        targetUrl: destAbsoluteTargetUrl
      }
    }), new CommandError('File Not Found.'));
  });

  it('correctly handles error when getCopyJobResult fails', async () => {
    spoUtilGetCopyJobResultStub.restore();
    spoUtilGetCopyJobResultStub = sinon.stub(spo, 'getCopyJobResult').rejects(new Error('Target file already exists.'));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl
      }
    }), new CommandError('Target file already exists.'));
  });
});
