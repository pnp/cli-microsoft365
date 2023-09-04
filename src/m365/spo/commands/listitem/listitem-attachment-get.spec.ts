import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
const command: Command = import('./listitem-attachment-get.js');

describe(commands.LISTITEM_ATTACHMENT_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '4fc5ba1e-18b7-49e0-81fe-54515cc2eede';
  const listTitle = 'Demo List';
  const listUrl = '/sites/project-x/Lists/DemoList';
  const listItemId = 147;
  const fileName = 'File1.jpg';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const attachmentResponse = {
    "FileName": "File1.jpg",
    "FileNameAsPath": {
      "DecodedUrl": "File1.jpg"
    },
    "ServerRelativePath": {
      "DecodedUrl": "/sites/project-x/Lists/DemoListAttachments/147/File1.jpg"
    },
    "ServerRelativeUrl": "/sites/project-x/Lists/DemoListAttachments/147/File1.jpg"
  };

  const getFakes = async (opts: any) => {
    if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
      return attachmentResponse;
    }

    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)/AttachmentFiles('${fileName}')`) {
      return attachmentResponse;
    }

    throw 'Invalid request';
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List', listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: listId, listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the specified listItemId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: 'a', fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('returns attachment from a list item by listId', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: webUrl,
      listId: listId,
      listItemId: listItemId,
      fileName: fileName
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentResponse));
  });

  it('returns attachment from a list item by listTitle', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: webUrl,
      listTitle: listTitle,
      listItemId: listItemId,
      fileName: fileName
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentResponse));
  });

  it('returns attachment from a list item by listUrl', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: webUrl,
      listUrl: listUrl,
      listItemId: listItemId,
      fileName: fileName
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    const options: any = {
      webUrl: webUrl,
      listId: listId,
      listItemId: listItemId,
      fileName: fileName
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles no attachment found', async () => {
    sinon.stub(request, 'get').rejects(new Error('Specified argument was out of the range of valid values.\r\nParameter name: fileName'));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listId: listId,
        listItemId: listItemId,
        fileName: fileName
      }
    } as any), new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: fileName'));
  });
});