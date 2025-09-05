import assert from 'assert';
import sinon from 'sinon';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import fs from 'fs';
import request from '../../../../request.js';
import command from './listitem-attachment-set.js';

describe(commands.LISTITEM_ATTACHMENT_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '236a0f92482d475bba8fd0e4f78555e4';
  const listTitle = 'Test list';
  const listUrl = 'sites/project-x/lists/testlist';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listItemId = 1;
  const filePath = 'C:\\Temp\\Test.pdf';
  const fileName = 'CLIRocks.pdf';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((_, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      request.put,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: 'invalid', listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and filePath exists', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listItemId option is not a valid number', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 'invalid', filePath: filePath, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('fails validation if the listId option is not a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', listItemId: listItemId, filePath: filePath, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if filePath does not exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('updates attachment to listitem in list retrieved by id while specifying fileName', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const putStub = sinon.stub(request, 'put').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles('${fileName}')/$value`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } });
    assert(putStub.called);
  });

  it('updates attachment to listitem in list retrieved by url while not specifying fileName', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const putStub = sinon.stub(request, 'put').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${listItemId})/AttachmentFiles('${fileName}')/$value`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, listUrl: listUrl, listItemId: listItemId, filePath: filePath, fileName: fileName } });
    assert(putStub.called);
  });

  it('updates attachment to listitem in list retrieved by url while specifying fileName without extension', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const fileNameWithoutExtension = fileName.split('.')[0];
    const fileNameWithExtension = `${fileNameWithoutExtension}.${filePath.split('.').pop()}`;
    const putStub = sinon.stub(request, 'put').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(${listItemId})/AttachmentFiles('${fileNameWithExtension}')/$value`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileNameWithoutExtension } });
    assert(putStub.called);
  });

  it('handles error when attachment does not exist', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('content read');
    const error = {
      error: {
        'odata.error': {
          code: '-2146233086, System.ArgumentOutOfRangeException',
          message: {
            lang: 'en-US',
            value: 'Specified argument was out of the range of valid values.\r\nParameter name: fileName'
          }
        }
      }
    };
    sinon.stub(request, 'put').callsFake(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles('${fileName}')/$value`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } }),
      new CommandError(error.error['odata.error'].message.value));
  });
});