import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command from './page-control-remove.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';
import { Page } from './Page.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.PAGE_CONTROL_REMOVE, () => {
  const spRootUrl = 'https://contoso.sharepoint.com';
  const serverRelWebUrl = '/sites/marketing';
  const webUrl = spRootUrl + serverRelWebUrl;
  const pageName = 'Home.aspx';
  const controlId = '12345678-1234-1234-1234-123456789012';
  const siteId = '4f67977c-3076-42e6-b3a6-5e7b73a76c67';
  const pageId = '969f1ecd-61a1-4577-a655-b8629fc8bfb9';

  const pageFileResponse = {
    UniqueId: pageId,
    SiteId: siteId
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let loggerLogSpy: sinon.SinonSpy;
  let confirmationPromptStub: sinon.SinonStub;
  let pagePublishStub: sinon.SinonStub;
  let deleteStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    pagePublishStub = sinon.stub(Page, 'publishPage').callsFake(async (webUrlParam: string, pageNameParam: string) => {
      if (webUrl === webUrlParam && pageName === pageNameParam) {
        return;
      }
      throw new Error(`Invalid parameters: webUrl=${webUrlParam}, pageName=${pageNameParam}`);
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(`${serverRelWebUrl}/SitePages/${pageName}`)}')?$select=UniqueId,SiteId`) {
        return pageFileResponse;
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `${spRootUrl}/_api/v2.1/sites/${pageFileResponse.SiteId}/pages/${pageFileResponse.UniqueId}/oneDrive.page/webParts/${controlId}`) {
        return;
      }

      throw 'Invalid DELETE request: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Page.publishPage,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_CONTROL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is invalid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', pageName: pageName, id: controlId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is invalid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, pageName: pageName, id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing the control', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the control when prompt is not confirmed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId } });
    assert(deleteStub.notCalled);
  });

  it('does not log a response', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly removes page control', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes page control without extension', async () => {
    const pageNameWithoutExtension = pageName.replace('.aspx', '');
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageNameWithoutExtension, id: controlId, force: true } });
    assert(deleteStub.calledOnce);
  });

  it('correctly removes page control and publishes page', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true, verbose: true } });
    assert(pagePublishStub.calledOnce);
  });

  it('correctly removes page control and keeps page as draft', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, draft: true, force: true } });
    assert(deleteStub.calledOnce);
    assert(pagePublishStub.notCalled);
  });

  it('correctly handles unexpected error', async () => {
    sinonUtil.restore(request.get);

    const errorMessage = 'The file /sites/marketing/SitePages/home.aspx does not exist.';
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-2130575338, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: errorMessage
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } }),
      new CommandError(errorMessage));
  });
});