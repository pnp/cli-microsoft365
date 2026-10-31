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
import command, { options } from './list-remove.js';

describe(commands.LIST_REMOVE, () => {
  const listId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const webUrl = 'https://contoso.sharepoint.com';
  const listTitle = 'Documents';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing list when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com', permanent: true }) });

    assert(promptIssued);
  });

  it('prompts before removing list when confirmation argument not passed (title)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ title: 'My list', webUrl: 'https://contoso.sharepoint.com' }) });

    assert(promptIssued);
  });

  it('aborts removing list when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com', permanent: true }) });
    assert(requests.length === 0);
  });

  it('removes the list when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `${webUrl}/_api/web/lists(guid'${listId}')`) {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com', permanent: true }) });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${webUrl}/_api/web/lists(guid'${listId}')`) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list when prompt confirmed by option', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `${webUrl}/_api/web/lists(guid'${listId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com', permanent: true, force: true }) });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `${webUrl}/_api/web/lists(guid'${listId}')`) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('uses correct recycle API URL by default', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${listTitle}')/recycle`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      })
    });

    assert(postStub.calledOnce);
  });

  it('command correctly handles list get reject request', async () => {
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
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${listTitle}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    const actionTitle: string = 'Documents';

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
        permanent: true,
        force: true
      })
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if both id and title options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Specify one of the following options: id, title.');
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '12345' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, '12345 is not a valid GUID');
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', title: 'Documents' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Specify one of the following options: id, title.');
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
