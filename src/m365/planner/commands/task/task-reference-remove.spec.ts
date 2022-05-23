import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-reference-remove');

describe(commands.TASK_REFERENCE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validUrl = 'https://www.microsoft.com';
  const validAlias = 'Test';

  const referenceResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word"
    }
  };

  const multiReferencesResponseNoAlias = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "type": "Word"
    },
    "https%3A//www%2Emicrosoft2%2Ecom": {
      "type": "Word"
    }
  };

  const multiReferencesResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word"
    },
    "https%3A//www%2Emicrosoft2%2Ecom": {
      "alias": "Test",
      "type": "Word"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    sinon.stub(Cli.getInstance().config, 'all').value({});
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
      cb({ continue: true });
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      Cli.getInstance().config.all
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_REFERENCE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if url does not contain http or https', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: 'www.microsoft.com'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid url with http specified', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: 'http://www.microsoft.com'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid url with https specified', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: 'https://www.microsoft.com'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets();
    assert.deepStrictEqual(optionSets, [['url', 'alias']]);
  });

  it('prompts before removal when confirm option not passed', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        url: validUrl
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

  it('passes validation when valid options specified', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly removes reference', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve({
          references: null
        });
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      taskId: validTaskId,
      url: validUrl,
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly removes reference by alias with prompting', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve({
          references: null
        });
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag",
          references: referenceResponse
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      taskId: validTaskId,
      alias: validAlias
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when task details endpoint fails', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('invalid')}/details`) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'invalid',
        url: validUrl
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Error fetching task details`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no references found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag",
          references: {}
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified reference with alias ${validAlias} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when reference does not contain alias', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag",
          references: multiReferencesResponseNoAlias
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified reference with alias ${validAlias} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple references found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag",
          references: multiReferencesResponse
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        alias: validAlias
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple references with alias ${validAlias} found. Pass one of the following urls within the "--url" option : https://www.microsoft.com,https://www.microsoft2.com`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});