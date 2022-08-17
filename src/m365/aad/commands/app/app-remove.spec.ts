import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let deleteRequestStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    promptOptions = undefined;
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        // fake call for getting app
        if (opts.url.indexOf('startswith') === -1) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications(id)",
            "value": [
              {
                "id": "d75be2e1-0204-4f95-857d-51a37cf40be8"
              }
            ]
          });
        }
      }
      return Promise.reject();
    });

    deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/d75be2e1-0204-4f95-857d-51a37cf40be8') {
        return Promise.resolve();
      }
      return Promise.reject();
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(commands.APP_DELETE) > -1), true);
  });

  it('fails validation if appId and objectId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', async () => {
    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId, nor name specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid guid', async () => {
    const actual = await command.validate({ options: { objectId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)', async () => {
    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: 'My app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the app when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the app when prompt not confirmed', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    }, () => {
      try {
        assert(deleteRequestStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes app when prompt confirmed (debug)', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes app with specified app (client) ID', (done) => {
    command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        confirm: true
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes app with specified object ID', (done) => {
    command.action(logger, {
      options: {
        objectId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        confirm: true
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes app with specified name', (done) => {
    command.action(logger, {
      options: {
        name: 'myapp',
        confirm: true
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get app by id when app does not exists', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject("No Azure AD application registration with ID myapp found");
    });

    command.action(logger, { options: { debug: true, appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8', confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("No Azure AD application registration with ID d75be2e1-0204-4f95-857d-51a37cf40be8 found")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get app by name when app does not exists', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject("No Azure AD application registration with name myapp found");
    });

    command.action(logger, { options: { debug: true, name: 'myapp', confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("No Azure AD application registration with name myapp found")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple apps with same name exists', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
          "value": [
            {
              "id": "d75be2e1-0204-4f95-857d-51a37cf40be8"
            },
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952"
            }
          ]
        });
      }

      return Promise.reject("Multiple Azure AD application registration with name myapp found. Please choose one of the object IDs: d75be2e1-0204-4f95-857d-51a37cf40be8, 340a4aa3-1af6-43ac-87d8-189819003952");
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'myapp',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Multiple Azure AD application registration with name myapp found. Please choose one of the object IDs: d75be2e1-0204-4f95-857d-51a37cf40be8, 340a4aa3-1af6-43ac-87d8-189819003952")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});