import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './page-control-remove.js';
import { Page } from './Page.js';

describe(commands.PAGE_CONTROL_REMOVE, () => {
  const spRootUrl = 'https://contoso.sharepoint.com';
  const serverRelWebUrl = '/sites/marketing';
  const webUrl = spRootUrl + serverRelWebUrl;
  const pageName = 'Home.aspx';
  const controlId = '12345678-1234-1234-1234-123456789012';

  const pageFileResponse = {
    CanvasContent1: `[{"position":{"layoutIndex":1,"zoneIndex":1,"zoneId":"fed1887c-145d-42fc-a459-ffc27f0cd3fa","sectionIndex":1,"sectionFactor":0,"controlIndex":1},"emphasis":{"zoneEmphasis":0},"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","controlType":3,"isFromSectionTemplate":false,"addedFromPersistedData":true,"webPartId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","reservedWidth":1536,"reservedHeight":228,"webPartData":{"id":"e69c4117-e349-45fd-9aba-eed00d741cfa","instanceId":"2cb8b5b5-f6dd-47ca-8740-8acc94e95714","title":"Banner","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.6","properties":{"title":"Experimental","imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","enableGradientEffect":true,"isDecorative":true,"isFullWidth":true,"authorByline":["i:0#.f|membership|admin@milanhdev.onmicrosoft.com"],"authors":[{"id":"i:0#.f|membership|admin@milanhdev.onmicrosoft.com","upn":"admin@milanhdev.onmicrosoft.com","email":"admin@milanhdev.onmicrosoft.com","name":"Milan Holemans","role":"Developer"}],"customContentDropSupport":"externallink","showTimeToRead":false},"containsDynamicDataSource":false}},{"position":{"layoutIndex":1,"zoneIndex":2,"zoneId":"f8e4f76b-3a1f-4569-8d7d-9f590553c6b4","sectionIndex":2,"sectionFactor":6,"controlIndex":1},"id":"${controlId}","controlType":3,"isFromSectionTemplate":false,"addedFromPersistedData":true,"webPartId":"0f087d7f-520e-42b7-89c0-496aaf979d58","reservedHeight":40,"reservedWidth":570,"webPartData":{"id":"0f087d7f-520e-42b7-89c0-496aaf979d58","instanceId":"1b7405d7-4aaf-49f7-82cb-2eca344a3c2c","title":"Button","description":"Add a clickable button with a custom label and link.","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{"label":"CLI for Microsoft 365"},"imageSources":{},"links":{"linkUrl":"https://pnp.github.io/cli-microsoft365/"}},"dataVersion":"1.1","properties":{"alignment":"Center","minimumLayoutWidth":5},"containsDynamicDataSource":false}},{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":false,"isSpellCheckEnabled":true,"globalRichTextStylingVersion":1,"rtePageSettings":{"contentVersion":5},"isEmailReady":false,"webPartsPageSettings":{"isTitleHeadingLevelsEnabled":false}}}]`
  };

  const removedCanvasContent1 = `[{"position":{"layoutIndex":1,"zoneIndex":1,"zoneId":"fed1887c-145d-42fc-a459-ffc27f0cd3fa","sectionIndex":1,"sectionFactor":0,"controlIndex":1},"emphasis":{"zoneEmphasis":0},"id":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","controlType":3,"isFromSectionTemplate":false,"addedFromPersistedData":true,"webPartId":"cbe7b0a9-3504-44dd-a3a3-0e5cacd07788","reservedWidth":1536,"reservedHeight":228,"webPartData":{"id":"e69c4117-e349-45fd-9aba-eed00d741cfa","instanceId":"2cb8b5b5-f6dd-47ca-8740-8acc94e95714","title":"Banner","audiences":[],"serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}},"dataVersion":"1.6","properties":{"title":"Experimental","imageSourceType":4,"layoutType":"FullWidthImage","textAlignment":"Left","showTopicHeader":false,"showPublishDate":false,"topicHeader":"","enableGradientEffect":true,"isDecorative":true,"isFullWidth":true,"authorByline":["i:0#.f|membership|admin@milanhdev.onmicrosoft.com"],"authors":[{"id":"i:0#.f|membership|admin@milanhdev.onmicrosoft.com","upn":"admin@milanhdev.onmicrosoft.com","email":"admin@milanhdev.onmicrosoft.com","name":"Milan Holemans","role":"Developer"}],"customContentDropSupport":"externallink","showTimeToRead":false},"containsDynamicDataSource":false}},{"controlType":0,"pageSettingsSlice":{"isDefaultDescription":true,"isDefaultThumbnail":false,"isSpellCheckEnabled":true,"globalRichTextStylingVersion":1,"rtePageSettings":{"contentVersion":5},"isEmailReady":false,"webPartsPageSettings":{"isTitleHeadingLevelsEnabled":false}}}]`;

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogSpy: sinon.SinonSpy;
  let confirmationPromptStub: sinon.SinonStub;
  let pagePublishStub: sinon.SinonStub;
  let patchStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      throw `Invalid parameters publishPage: webUrl=${webUrlParam}, pageName=${pageNameParam}`;
    });

    sinon.stub(Page, 'checkout').callsFake(async (pageNameParam: string, webUrlParam: string) => {
      if (webUrl === webUrlParam && pageName === pageNameParam) {
        return pageFileResponse as any;
      }

      throw `Invalid parameters checkout: webUrl=${webUrlParam}, pageName=${pageNameParam}`;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')?$select=CanvasContent1`) {
        return pageFileResponse;
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')/SavePageAsDraft`) {
        return;
      }

      throw 'Invalid PATCH request: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      Page.publishPage,
      Page.checkout,
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
    assert(patchStub.notCalled);
  });

  it('does not log a response', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly removes page control', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } });

    assert.strictEqual(patchStub.firstCall.args[0].data.CanvasContent1, removedCanvasContent1);
  });

  it('correctly removes page control without extension', async () => {
    const pageNameWithoutExtension = pageName.replace('.aspx', '');
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageNameWithoutExtension, id: controlId, force: true } });

    assert.strictEqual(patchStub.firstCall.args[0].data.CanvasContent1, removedCanvasContent1);
  });

  it('correctly removes page control and publishes page', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true, verbose: true } });

    assert(pagePublishStub.calledOnce);
  });

  it('correctly removes page control and keeps page as draft', async () => {
    await command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, draft: true, force: true } });

    assert.strictEqual(patchStub.firstCall.args[0].data.CanvasContent1, removedCanvasContent1);
    assert(pagePublishStub.notCalled);
  });

  it('correctly handles error when page does not contain canvas controls', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').resolves({ CanvasContent1: '' });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: controlId, force: true } }),
      new CommandError(`Page '${pageName}' doesn't contain canvas control '${controlId}'.`));
  });

  it('correctly handles error when control with specified ID is not found on page', async () => {
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, pageName: pageName, id: 'b27bbb8f-8aa2-4bd8-b11b-5addf65983b0', force: true } }),
      new CommandError(`Control with ID 'b27bbb8f-8aa2-4bd8-b11b-5addf65983b0' was not found on page '${pageName}'.`));
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