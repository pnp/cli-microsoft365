import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './contenttype-remove.js';

describe(commands.CONTENTTYPE_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    promptOptions = undefined;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('delete content type by id - prompt', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', force: false }
    } as any);
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });


  it('delete content type by id - prompt:continue', async () => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        force: false
      }
    } as any);
    assert(postCallbackStub.called);
  });


  it('delete content type by id - prompt:declined', async () => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        force: false
      }
    } as any);
    assert(postCallbackStub.notCalled);
  });

  it('delete content type by name - prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return { "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', force: false } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });


  it('delete content type by name - prompt:continue', async () => {
    const getCallbackStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return { "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] };
      }

      throw 'Invalid request';
    });

    const postCallbackStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, verbose: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', force: false } });
    assert(getCallbackStub.called);
    assert(postCallbackStub.called);
  });

  it('delete content type by name - prompt:declined', async () => {
    const postCallbackStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return { "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', force: false } });
    assert(postCallbackStub.notCalled);
  });

  it('correctly escapes special characters in the content type name', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'Test%20Content%20Type')`) > -1) {
        return { "value": [{ "Name": "Test Content Type", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] };
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'Test Content Type', force: true } } as any);
    assert(getStub.called);
    assert(postStub.called);
  });


  it('correctly handles site content type not found by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', force: true } } as any),
      new CommandError('Content type not found'));
  });

  it('correctly handles site content type not found by name', async () => {
    //NonExistentContentType
    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'NonExistentContentType')`) > -1) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    const deleteRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes`) > -1) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', force: true } } as any),
      new CommandError('Content type not found'));

    assert(getRequestStub.called);
    assert(deleteRequestStub.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = command.types;
    ['i', 'id'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('supports verbose mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the content type ID nor content type Name parameters are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when contenttype id parameter is provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name parameter is provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype id and confirm parameters are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name and confirm parameters are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when neither name nor id are provided, but confirm is', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both name and id are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
