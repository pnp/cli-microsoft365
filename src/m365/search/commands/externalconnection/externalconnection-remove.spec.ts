import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-remove');

describe(commands.EXTERNALCONNECTION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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

    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'name']]);
  });

  it('prompts before removing the specified external connection by id when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        id: "contosohr"
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

  it('prompts before removing the specified external connection by name when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        name: "Contoso HR"
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

  it('aborts removing the specified external connection when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    command.action(logger, { options: { debug: true, id: "contosohr" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified external connection when prompt confirmed (debug)', (done) => {
    let externalConnectionRemoveCallIssued = false;

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        externalConnectionRemoveCallIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, id: "contosohr" } }, () => {
      try {
        assert(externalConnectionRemoveCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified external connection without prompting when confirm specified', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: "contosohr", confirm: true } }, () => {
      done();
    });
  });

  it('removes external connection with specified ID', (done) => {
    sinon.stub(request, 'delete').callsFake((opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return Promise.resolve();
      }
      return Promise.reject();
    });

    command.action(logger, { options: { debug: false, id: "contosohr", confirm: true } }, () => {
      done();
    });
  });

  it('removes external connection with specified name', (done) => {
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq `) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "contosohr",
              "name": "Contoso HR",
              "description": "Connection to index Contoso HR system"
            }
          ]
        });
      }
      return Promise.reject();
    });

    sinon.stub(request, 'delete').callsFake((opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return Promise.resolve();
      }
      return Promise.reject();
    });

    command.action(logger, { options: { debug: false, name: "Contoso HR", confirm: true } }, () => {
      done();
    });
  });

  it('fails to get external connection by name when it does not exists', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('The specified connection does not exist in Microsoft Search');
    });

    command.action(logger, { options: { debug: false, name: "Fabrikam HR", confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("The specified connection does not exist in Microsoft Search")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple external connections with same name exists', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "fabrikamhr"
            },
            {
              "id": "contosohr"
            }
          ]
        });
      }

      return Promise.reject("Invalid request");
    });

    command.action(logger, {
      options: {
        debug: false,
        name: "My HR",
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Multiple external connections with name My HR found. Please disambiguate (IDs): fabrikamhr, contosohr")));
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