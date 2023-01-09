import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
const command: Command = require('./file-retentionlabel-ensure');

describe(commands.FILE_RETENTIONLABEL_ENSURE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const fileUrl = `/Shared Documents/Fo'lde'r/Document.docx`;
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const retentionlabelName = "retentionlabel";
  const SpoListItemRetentionLabelEnsureCommandOutput = `{ "stdout": "", "stderr": "" }`;

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_RETENTIONLABEL_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds the retentionlabel from a file based on fileUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileUrl)}')?$expand=ListItemAllFields`) {
        return { ListItemAllFields: { Id: listId }, ServerRelativeUrl: fileUrl };
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: false,
        fileUrl: fileUrl,
        webUrl: webUrl,
        name: retentionlabelName
      }
    });
    assert(postSpy.called);
  });

  it('adds the retentionlabel from a file based on fileId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileById('${fileId}')?$expand=ListItemAllFields`) {
        return { ListItemAllFields: { Id: listId }, ServerRelativeUrl: fileUrl };
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
        debug: false,
        fileId: fileId,
        webUrl: webUrl,
        name: retentionlabelName
      }
    });
    assert(postSpy.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'get').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: retentionlabelName,
        fileUrl: fileUrl,
        webUrl: webUrl
      }
    }), new CommandError(errorMessage));
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

  it('fails validation if both fileUrl or fileId options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '12345', name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fileId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both fileId and fileUrl options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});