import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './listitem-attachment-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LISTITEM_ATTACHMENT_LIST, () => {
  const listUrl = 'sites/project-x/documents';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const attachmentsResponse = {
    AttachmentFiles: [
      {
        "FileName": "my_file.docx",
        "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_file.docx"
      },
      {
        "FileName": "my_workbook.xlsx",
        "ServerRelativeUrl": "/sites/project-x/Lists/Demo List/Attachments/1/my_workbook.xlsx"
      }
    ]
  };

  const listItemId = 147;

  const getFakes = async (opts: any) => {
    if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
      return attachmentsResponse;
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)?$select=AttachmentFiles&$expand=AttachmentFiles`) {
      return attachmentsResponse;
    }
    throw 'Invalid request';
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ATTACHMENT_LIST), true);
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

  it('fails validation if listTitle and listId option not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listItemId: listItemId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: listItemId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listTitle: 'Demo List', listItemId: listItemId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: listItemId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', listItemId: listItemId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listItemId: listItemId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the specified listItemId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['FileName', 'ServerRelativeUrl']);
  });

  it('returns attachments associated to a list item by listId', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      listItemId: listItemId
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentsResponse.AttachmentFiles));
  });

  it('returns attachments associated to a list item by listTitle', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listTitle: 'Demo List',
      listItemId: listItemId
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentsResponse.AttachmentFiles));
  });

  it('returns attachments associated to a list item by listUrl', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);

    const options: any = {
      verbose: true,
      webUrl: webUrl,
      listUrl: listUrl,
      listItemId: listItemId
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(attachmentsResponse.AttachmentFiles));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    const options: any = {
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      listItemId: listItemId,
      output: "json"
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles No attachments found (debug)', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.resolve({ AttachmentFiles: [] });
    });

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      listItemId: listItemId
    };

    await command.action(logger, { options: options });
    assert(loggerLogToStderrSpy.calledWith('No attachments found'));
  });
});
