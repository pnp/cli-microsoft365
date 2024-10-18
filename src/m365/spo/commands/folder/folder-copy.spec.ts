import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './folder-copy.js';
import { CreateFolderCopyJobsNameConflictBehavior, spo } from '../../../../utils/spo.js';
import { settingsNames } from '../../../../settingsNames.js';
import { CommandError } from '../../../../Command.js';

const sourceWebUrl = 'https://contoso.sharepoint.com/sites/Sales';
const sourceFolderName = 'Logos';
const sourceServerRelUrl = '/sites/Sales/Shared Documents/' + sourceFolderName;
const sourceSiteRelUrl = '/Shared Documents/' + sourceFolderName;
const sourceAbsoluteUrl = 'https://contoso.sharepoint.com' + sourceServerRelUrl;
const sourceFolderId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';

const destWebUrl = 'https://contoso.sharepoint.com/sites/Marketing';
const destSiteRelUrl = '/Documents';
const destServerRelUrl = '/sites/Marketing' + destSiteRelUrl;
const destAbsoluteTargetUrl = 'https://contoso.sharepoint.com' + destServerRelUrl;
const destFolderId = '15488d89-b82b-40be-958a-922b2ed79383';

const copyJobInfo = {
  EncryptionKey: '2by8+2oizihYOFqk02Tlokj8lWUShePAEE+WMuA9lzA=',
  JobId: 'd812e5a0-d95a-4e4f-bcb7-d4415e88c8ee',
  JobQueueUri: 'https://spoam1db1m020p4.queue.core.windows.net/2-1499-20240831-29533e6c72c6464780b756c71ea3fe92?sv=2018-03-28&sig=aX%2BNOkUimZ3f%2B%2BvdXI95%2FKJI1e5UE6TU703Dw3Eb5c8%3D&st=2024-08-09T00%3A00%3A00Z&se=2024-08-31T00%3A00%3A00Z&sp=rap',
  SourceListItemUniqueIds: [
    sourceFolderId
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
  TargetObjectUniqueId: destFolderId,
  TargetObjectSiteRelativeUrl: destSiteRelUrl.substring(1),
  CorrelationId: '5efd44a1-c034-9000-9692-4e1a1b3ca33b'
};

const destFolderResponse = {
  Exists: true,
  ExistsAllowThrowForPolicyFailures: true,
  ExistsWithException: true,
  IsWOPIEnabled: false,
  ItemCount: 6,
  Name: sourceFolderName,
  ProgID: null,
  ServerRelativeUrl: destServerRelUrl,
  TimeCreated: '2024-09-26T20:52:07Z',
  TimeLastModified: '2024-09-26T21:16:26Z',
  UniqueId: '59abed95-34f9-470b-a133-ae8932480b53',
  WelcomePage: ''
};

describe(commands.FOLDER_COPY, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;

  let spoUtilCreateCopyJobStub: sinon.SinonStub;
  let spoUtilGetCopyJobResultStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => settingName === settingsNames.prompt ? false : defaultValue);
    spoUtilCreateCopyJobStub = sinon.stub(spo, 'createFolderCopyJob').resolves(copyJobInfo);
    spoUtilGetCopyJobResultStub = sinon.stub(spo, 'getCopyJobResult').resolves(copyJobResult);
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
    spoUtilGetCopyJobResultStub.resetHistory();
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['targetUrl', 'sourceUrl']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: sourceServerRelUrl, targetUrl: destServerRelUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sourceId is not a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceId: 'invalid', targetUrl: destServerRelUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if nameConflictBehavior is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceUrl: sourceServerRelUrl, targetUrl: destServerRelUrl, nameConflictBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the sourceId is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceId: sourceFolderId, targetUrl: destServerRelUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: sourceWebUrl, sourceUrl: sourceServerRelUrl, targetUrl: destServerRelUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly outputs exactly one result when folder is copied when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFolderById('${sourceFolderId}')/ServerRelativePath`) {
        return {
          DecodedUrl: destAbsoluteTargetUrl + `/${sourceFolderName}`
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceId: sourceFolderId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert(loggerLogSpy.calledOnce);
  });

  it('correctly outputs result when folder is copied when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFolderById('${sourceFolderId}')/ServerRelativePath`) {
        return {
          DecodedUrl: sourceAbsoluteUrl
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: sourceWebUrl,
        sourceId: sourceFolderId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert(loggerLogSpy.calledOnce);
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], destFolderResponse);
  });

  it('correctly outputs result when folder is copied when using sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
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
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], destFolderResponse);
  });

  it('correctly copies a folder when using sourceId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${sourceWebUrl}/_api/Web/GetFolderById('${sourceFolderId}')/ServerRelativePath`) {
        return {
          DecodedUrl: sourceAbsoluteUrl
        };
      }

      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceId: sourceFolderId,
        targetUrl: destAbsoluteTargetUrl
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Fail,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a folder when using sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
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
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Fail,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a folder when using site-relative sourceUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceSiteRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'fail'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Fail,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a folder when using absolute urls', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceAbsoluteUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'rename'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        operation: 'copy',
        newName: undefined
      }
    ]);
  });

  it('correctly copies a folder when using sourceUrl with extra options', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl,
        nameConflictBehavior: 'rename',
        newName: 'Folder-renamed'
      }
    });

    assert.deepStrictEqual(spoUtilCreateCopyJobStub.lastCall.args, [
      sourceWebUrl,
      sourceAbsoluteUrl,
      destAbsoluteTargetUrl,
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        operation: 'copy',
        newName: 'Folder-renamed'
      }
    ]);
  });

  it('correctly polls for the copy job to finish', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${destWebUrl}/_api/Web/GetFolderById('${destFolderId}')`) {
        return destFolderResponse;
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
          message: {
            lang: 'en-US',
            value: 'Folder Not Found.'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceId: sourceFolderId,
        targetUrl: destAbsoluteTargetUrl
      }
    }), new CommandError('Folder Not Found.'));
  });

  it('correctly handles error when getCopyJobResult fails', async () => {
    spoUtilGetCopyJobResultStub.restore();
    spoUtilGetCopyJobResultStub = sinon.stub(spo, 'getCopyJobResult').rejects(new Error('Target folder already exists.'));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: sourceWebUrl,
        sourceUrl: sourceServerRelUrl,
        targetUrl: destAbsoluteTargetUrl
      }
    }), new CommandError('Target folder already exists.'));
  });
});