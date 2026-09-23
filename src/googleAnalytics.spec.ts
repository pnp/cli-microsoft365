import assert from 'assert';
import Axios from 'axios';
import Configstore from 'configstore';
import fs from 'fs';
import sinon from 'sinon';
import { googleAnalytics } from './googleAnalytics.js';

const env = { ...process.env };

describe('googleAnalytics', () => {
  let axiosPostStub: sinon.SinonStub;

  beforeEach(() => {
    sinon.stub(Configstore.prototype, 'get').returns('client-id');
    axiosPostStub = sinon.stub(Axios, 'post').resolves({ status: 204 });
  });

  afterEach(() => {
    sinon.restore();
    process.env = env;
  });

  it('sends command and option usage using Measurement Protocol', async () => {
    await googleAnalytics.trackEvent('spo file add', { output: 'json', debug: false }, {
      sessionId: 'session-id',
      shell: 'zsh'
    });

    assert.strictEqual(axiosPostStub.callCount, 1);
    assert.match(axiosPostStub.firstCall.args[0], /measurement_id=.+&api_secret=.+/);

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.client_id, 'client-id');
    assert.strictEqual(body.events.length, 1);
    assert.strictEqual(body.events[0].name, 'command_used');
    assert.strictEqual(body.events[0].params.command_name, 'spo file add');
    assert.strictEqual(body.events[0].params.shell, 'zsh');
    assert.strictEqual(body.events[0].params.os, process.platform);
    assert.strictEqual(body.events[0].params.output, 'json');
    assert.strictEqual(body.events[0].params.debug, 'false');

    const requestOptions = axiosPostStub.firstCall.args[2];
    assert.strictEqual(requestOptions.timeout, 3000);
    assert.strictEqual(requestOptions.validateStatus(500), true);
  });

  it('uses the package version outside of development', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const modulePath = './googleAnalytics.js?production';
    const { googleAnalytics: productionGoogleAnalytics } = await import(modulePath);

    assert(!productionGoogleAnalytics.commonProperties.version.endsWith('-dev'));
  });

  it('creates and stores a client ID when one does not exist', async () => {
    (Configstore.prototype.get as sinon.SinonStub).returns(undefined);
    const setStub = sinon.stub(Configstore.prototype, 'set');

    await googleAnalytics.trackEvent('help', {}, { sessionId: 'session-id', shell: 'zsh' });

    assert(setStub.calledOnce);
    const args = setStub.firstCall.args as any[];
    assert.strictEqual(args[0], 'telemetryClientId');
    assert.match(args[1], /^[0-9a-f-]{36}$/);
  });

  it('omits undefined option properties', async () => {
    await googleAnalytics.trackEvent('test command', { output: 'json', query: undefined }, { sessionId: 'session-id', shell: 'zsh' });

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.events.length, 1);
    assert.strictEqual(body.events[0].params.output, 'json');
    assert.strictEqual(body.events[0].params.query, undefined);
  });

  it('converts non-primitive option properties to strings', async () => {
    await googleAnalytics.trackEvent('test command', { value: null }, { sessionId: 'session-id', shell: 'zsh' });

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.events[0].params.value, 'null');
  });

  it('sends numeric option properties as numbers', async () => {
    await googleAnalytics.trackEvent('test command', { value: 1 }, { sessionId: 'session-id', shell: 'zsh' });

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.events[0].params.value, 1);
  });

  it('sets the Docker environment in telemetry', async () => {
    process.env.CLIMICROSOFT365_ENV = 'docker';
    const modulePath = `./googleAnalytics.js?docker=${Math.random()}`;
    const { googleAnalytics: dockerGoogleAnalytics } = await import(modulePath);

    assert.strictEqual(dockerGoogleAnalytics.commonProperties.env, 'docker');
  });

  it('sends errors as exceptions using Measurement Protocol', async () => {
    const message = 'a'.repeat(101);

    await googleAnalytics.trackException(new Error(message), { sessionId: 'session-id', shell: 'zsh' });

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.events[0].name, 'exception');
    assert.strictEqual(body.events[0].params.description, message.substring(0, 100));
    assert.strictEqual(body.events[0].params.shell, 'zsh');
  });

  it('converts non-error exceptions to strings', async () => {
    await googleAnalytics.trackException('error', { sessionId: 'session-id', shell: 'zsh' });

    const body = axiosPostStub.firstCall.args[1];
    assert.strictEqual(body.events[0].params.description, 'error');
  });

  it('throws when Google Analytics rejects the request', async () => {
    axiosPostStub.resolves({ status: 500 });

    await assert.rejects(
      googleAnalytics.trackEvent('test command', {}, { sessionId: 'session-id', shell: 'zsh' }),
      /Google Analytics returned 500/
    );
  });

  it('throws when Google Analytics returns a status below 200', async () => {
    axiosPostStub.resolves({ status: 199 });

    await assert.rejects(
      googleAnalytics.trackEvent('test command', {}, { sessionId: 'session-id', shell: 'zsh' }),
      /Google Analytics returned 199/
    );
  });
});