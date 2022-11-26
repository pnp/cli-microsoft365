import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
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
const command: Command = require('./contenttype-set');

describe(commands.CONTENTTYPE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '00000000-0000-0000-0000-000000000000';
  const listTitle = 'Assets';
  const listUrl = '/sites/project-x/Lists/Assets';
  const id = '0x0101';
  const name = 'Asset';
  const newName = 'New asset name';

  const contentTypesResponse = {
    value: [
      {
        Name: name,
        Group: 'Custom group',
        Id: {
          StringValue: id
        }
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name'] }]);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is specified and is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId and listTitle is specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listId: listId, listTitle: listTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId and listUrl is specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listId: listId, listUrl: listUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listUrl is specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listTitle: listTitle, listUrl: listUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId, listUrl and is specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listId: listId, listUrl: listUrl, listTitle: listTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl, id and listId are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, listId: listId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when webUrl, name are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, name: name } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('correctly updates content type with id', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await command.action(logger, { options: { webUrl: webUrl, id: id, Name: newName } } as any);
    assert.strictEqual(patchStub.lastCall.args[0].url, `${webUrl}/_api/Web/ContentTypes/GetById('${id}')`);
  });

  it('correctly updates content type with name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/Web/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return Promise.resolve(contentTypesResponse);
      }

      return Promise.reject('Invalid request url: ' + opts.url);
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await command.action(logger, { options: { webUrl: webUrl, name: name, Name: newName } } as any);
    assert.strictEqual(patchStub.lastCall.args[0].url, `${webUrl}/_api/Web/ContentTypes/GetById('${contentTypesResponse.value[0].Id.StringValue}')`);
  });

  it('correctly updates content type with name and listId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/Web/Lists/GetById('${formatting.encodeQueryParameter(listId)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return Promise.resolve(contentTypesResponse);
      }

      return Promise.reject('Invalid request url: ' + opts.url);
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await command.action(logger, { options: { webUrl: webUrl, name: name, listId: listId, Name: newName } } as any);
    assert.strictEqual(patchStub.lastCall.args[0].url, `${webUrl}/_api/Web/Lists/GetById('${formatting.encodeQueryParameter(listId)}')/ContentTypes/GetById('${contentTypesResponse.value[0].Id.StringValue}')`);
  });

  it('correctly updates content type with name and listTitle', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/Web/Lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return Promise.resolve(contentTypesResponse);
      }

      return Promise.reject('Invalid request url: ' + opts.url);
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await command.action(logger, { options: { webUrl: webUrl, name: name, listTitle: listTitle, Name: newName } } as any);
    assert.strictEqual(patchStub.lastCall.args[0].url, `${webUrl}/_api/Web/Lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/ContentTypes/GetById('${contentTypesResponse.value[0].Id.StringValue}')`);
  });

  it('correctly updates content type with name and listUrl', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return Promise.resolve(contentTypesResponse);
      }

      return Promise.reject('Invalid request url: ' + opts.url);
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await command.action(logger, { options: { webUrl: webUrl, name: name, listUrl: listUrl, Name: newName } } as any);
    assert.strictEqual(patchStub.lastCall.args[0].url, `${webUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')/ContentTypes/GetById('${contentTypesResponse.value[0].Id.StringValue}')`);
  });

  it('fails to update content type with name and listUrl when content type does not exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `${webUrl}/_api/Web/ContentTypes?$filter=Name eq '${name}'&$select=Id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request url: ' + opts.url);
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(() => Promise.resolve());

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, name: name, Name: newName } } as any), new CommandError(`The specified content type '${name}' does not exist`));
    assert(patchStub.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'patch').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id, Name: newName } } as any), new CommandError('An error has occurred'));
  });
});