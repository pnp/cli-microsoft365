import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import command, { options } from './list-sensitivitylabel-remove.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';

describe(commands.LIST_SENSITIVITYLABEL_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listTitle = 'Shared Documents';
  const listId = 'b4cfa0d9-b3d7-49ae-a0f0-f14ffdd005f7';
  const listUrl = '/Shared Documents';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_SENSITIVITYLABEL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when listTitle is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: listTitle });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when listId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when listUrl is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listUrl: listUrl });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if webUrl is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listTitle: listTitle });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if listId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', () => {
    const schema = commandInfo.command.getRefinedSchema(commandOptionsSchema);
    const actual = schema!.safeParse({ webUrl: webUrl });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if multiple list options are specified', () => {
    const schema = commandInfo.command.getRefinedSchema(commandOptionsSchema);
    const actual = schema!.safeParse({ webUrl: webUrl, listTitle: listTitle, listId: listId });
    assert.notStrictEqual(actual.success, true);
  });

  it('prompts before removing the sensitivity label when force option not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle } });

    assert(promptIssued);
  });

  it('prompts before removing the sensitivity label when using listUrl and force option not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, listUrl: listUrl } });

    assert(promptIssued);
  });

  it('aborts removing sensitivity label when prompt not confirmed', async () => {
    const patchSpy = sinon.stub(request, 'patch').resolves();

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle } });
    assert(patchSpy.notCalled);
  });

  it('removes sensitivity label from document library using listTitle without prompting for confirmation', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, force: true, verbose: true } });

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, '');
  });

  it('removes sensitivity label from document library using listId without prompting for confirmation', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'b4cfa0d9-b3d7-49ae-a0f0-f14ffdd005f7')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listId: listId, force: true, verbose: true } });

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, '');
  });

  it('removes sensitivity label from document library using listUrl without prompting for confirmation', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList('%2FShared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listUrl: listUrl, force: true, verbose: true } });

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, '');
  });

  it('removes sensitivity label when prompt is confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle } });

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, '');
  });

  it('correctly handles API error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: '404 - File not found'
          }
        }
      }
    };

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
