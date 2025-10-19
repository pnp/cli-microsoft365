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
import command from './field-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FIELD_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

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

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;

    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FIELD_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing field when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } });

    assert(promptIssued);
  });

  it('prompts before removing field when confirmation argument not passed (title)', async () => {
    await command.action(logger, { options: { title: 'myfield1', webUrl: 'https://contoso.sharepoint.com' } });

    assert(promptIssued);
  });

  it('prompts before removing list column when confirmation argument not passed', async () => {
    await command.action(logger, { options: { title: 'myfield1', webUrl: 'https://contoso.sharepoint.com', listTitle: 'My List' } });

    assert(promptIssued);
  });

  it('prompts before removing field when confirmation argument not passed (internalName)', async () => {
    await command.action(logger, { options: { internalName: 'myfield1', webUrl: 'https://contoso.sharepoint.com' } });

    assert(promptIssued);
  });

  it('prompts before removing list column when confirmation argument not passed (internalName)', async () => {
    await command.action(logger, { options: { internalName: 'myfield1', webUrl: 'https://contoso.sharepoint.com', listTitle: 'My List' } });

    assert(promptIssued);
  });

  it('aborts removing field when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } });
    assert(requests.length === 0);
  });

  it('aborts removing field when prompt not confirmed and passing the group parameter', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { group: 'MyGroup', webUrl: 'https://contoso.sharepoint.com' } });
    assert(requests.length === 0);
  });

  it('removes the field when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields(guid'b2307a39-e878-458b-bc90-03bc578531d6')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } }));
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/fields/getbyid('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('command correctly handles field get reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyinternalnameortitle('field1')`) {
        throw err;
      }

      throw 'Invalid request';
    });

    const actionTitle: string = 'field1';

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    }), new CommandError(err));
  });

  it('calls the correct remove url when id and list url specified', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList('%2Fsites%2Fportal%2FLists%2FEvents')/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listUrl: 'Lists/Events', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
  });

  it('calls group and deletes two fields and asks for confirmation', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList('%2Fsites%2Fportal%2FLists%2FEvents')/fields`) {
        return {
          'value': [{
            'Id': '03e45e84-1992-4d42-9116-26f756012634',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012635',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012636',
            'Group': 'DifferentGroup'
          }]
        };
      }
      throw 'Invalid request';
    });

    const deletion = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList('%2Fsites%2Fportal%2FLists%2FEvents')/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList('%2Fsites%2Fportal%2FLists%2FEvents')/fields/getbyid('03e45e84-1992-4d42-9116-26f756012635')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012635'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', listUrl: '/sites/portal/Lists/Events' } });
    assert(getStub.called);
    assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
    assert.strictEqual(deletion.secondCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')');
    assert.strictEqual(deletion.callCount, 2);
  });

  it('calls group and deletes two fields', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return {
          'value': [{
            'Id': '03e45e84-1992-4d42-9116-26f756012634',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012635',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012636',
            'Group': 'DifferentGroup'
          }]
        };
      }
      throw 'Invalid request';
    });

    const deletion = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012635')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012635'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', force: true } });
    assert(getStub.called);
    assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
    assert.strictEqual(deletion.secondCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')');
    assert.strictEqual(deletion.callCount, 2);
  });

  it('calls group and deletes no fields', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return {
          'value': [{
            'Id': '03e45e84-1992-4d42-9116-26f756012634',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012635',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012636',
            'Group': 'DifferentGroup'
          }]
        };
      }
      throw 'Invalid request';
    });

    const deletion = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields`) > -1) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup1', force: true } });
    assert(getStub.called);
    assert(deletion.notCalled);
  });

  it('handles failure when get operation fails', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    };

    const getStub = sinon.stub(request, 'get').rejects(error);

    const deletion = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012635')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012635'
        };
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        throw error;
      }

      throw error;
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', force: true } } as any),
      new CommandError('Invalid request'));
    assert(getStub.called);
    assert(deletion.notCalled);
  });

  it('handles failure when one deletion fails', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    };
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return {
          'value': [{
            'Id': '03e45e84-1992-4d42-9116-26f756012634',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012635',
            'Group': 'MyGroup'
          },
          {
            'Id': '03e45e84-1992-4d42-9116-26f756012636',
            'Group': 'DifferentGroup'
          }]
        };
      }
      throw 'Invalid request';
    });

    const deletion = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012635')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012635'
        };
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        throw error;
      }

      throw error;
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', force: true } } as any),
      new CommandError(error.error['odata.error'].message.value));
    assert(getStub.called);
    assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
    assert.strictEqual(deletion.callCount, 2);
  });

  it('calls the correct get url when field title and list title specified (verbose)', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'Title', listTitle: 'Documents', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('calls the correct get url when field title and list title specified', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'Title', listTitle: 'Documents', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('calls the correct get url when field title and list url specified', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid'03e45e84-1992-4d42-9116-26f756012634')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', title: 'Title', listId: '03e45e84-1992-4d42-9116-26f756012634', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid\'03e45e84-1992-4d42-9116-26f756012634\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('calls the correct get url when field internalName and list title specified (verbose)', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', internalName: 'Title', listTitle: 'Documents', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('calls the correct get url when field internalName and list title specified', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', internalName: 'Title', listTitle: 'Documents', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('calls the correct get url when field internalName and list url specified', async () => {
    const getStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid'03e45e84-1992-4d42-9116-26f756012634')/fields/getbyinternalnameortitle('Title')`) {
        return {
          'Id': '03e45e84-1992-4d42-9116-26f756012634'
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', internalName: 'Title', listId: '03e45e84-1992-4d42-9116-26f756012634', force: true } });
    assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid\'03e45e84-1992-4d42-9116-26f756012634\')/fields/getbyinternalnameortitle(\'Title\')');
  });

  it('correctly handles site column not found', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyinternalnameortitle('actionTitle')`) {
        throw error;
      }
      throw 'Invalid request';
    });
    const actionTitle: string = 'field1';

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', title: actionTitle, force: true } } as any),
      new CommandError(error.error['odata.error'].message.value));
  });

  it('correctly handles list column not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        throw {
          error: {
            'odata.error': {
              'code': '-2147024809, System.ArgumentException',
              'message': {
                'lang': 'en-US',
                'value': 'Invalid field name. {03e45e84-1992-4d42-9116-26f756012634}  /sites/portal/Shared Documents'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listTitle: 'Documents', force: true } } as any),
      new CommandError('Invalid field name. {03e45e84-1992-4d42-9116-26f756012634}  /sites/portal/Shared Documents'));
  });

  it('correctly handles list not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/fields/getbyid('03e45e84-1992-4d42-9116-26f756012634')`) {
        throw {
          error: {
            'odata.error': {
              'code': '-1, System.ArgumentException',
              'message': {
                'lang': 'en-US',
                'value': 'List \'Documents\' does not exist at site with URL \'https://contoso.sharepoint.com/sites/portal\'.'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listTitle: 'Documents', force: true } } as any),
      new CommandError('List \'Documents\' does not exist at site with URL \'https://contoso.sharepoint.com/sites/portal\'.'));
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

  it('fails validation if either of id, title and internalName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', force: true, listTitle: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the field ID option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the field ID option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the list ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 'BC448D63-484F-49C5-AB8C-96B14AA68D50',
        force: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
