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

    assert.strictEqual(axiosPostStub.callCount, 3);
    assert.match(axiosPostStub.firstCall.args[0], /measurement_id=.+&api_secret=.+/);

    const commandBody = axiosPostStub.firstCall.args[1];
    assert.strictEqual(commandBody.client_id, 'client-id');
    assert.strictEqual(commandBody.events[0].name, 'command_used');
    assert.strictEqual(commandBody.events[0].params.command_name, 'spo file add');
    assert.strictEqual(commandBody.events[0].params.shell, 'zsh');
    assert.strictEqual(commandBody.events[0].params.os, process.platform);

    const outputBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(outputBody.events[0].name, 'command_option_used');
    assert.strictEqual(outputBody.events[0].params.option_name, 'output');
    assert.strictEqual(outputBody.events[0].params.option_value, 'json');

    const debugBody = axiosPostStub.thirdCall.args[1];
    assert.strictEqual(debugBody.events[0].params.option_name, 'debug');
    assert.strictEqual(debugBody.events[0].params.option_value, 'false');

    const requestOptions = axiosPostStub.firstCall.args[2];
    assert.strictEqual(requestOptions.timeout, 1000);
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

    assert.strictEqual(axiosPostStub.callCount, 2);
    const outputBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(outputBody.events[0].params.option_name, 'output');
    assert.strictEqual(outputBody.events[0].params.option_value, 'json');
  });

  it('converts non-primitive option properties to strings', async () => {
    await googleAnalytics.trackEvent('test command', { value: null }, { sessionId: 'session-id', shell: 'zsh' });

    const optionBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(optionBody.events[0].params.option_value, 'null');
  });

  it('sends numeric option properties as numbers', async () => {
    await googleAnalytics.trackEvent('test command', { value: 1 }, { sessionId: 'session-id', shell: 'zsh' });

    const optionBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(optionBody.events[0].params.option_value, 1);
  });

  it('sends option names longer than 40 characters as parameter values', async () => {
    const optionName = 'disableTrialEnvironmentCreationByNonAdminUsers';

    await googleAnalytics.trackEvent('test command', { [optionName]: true }, { sessionId: 'session-id', shell: 'zsh' });

    const optionBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(optionBody.events[0].params.option_name, optionName);
    assert.strictEqual(optionBody.events[0].params.option_value, 'true');
  });

  it('shortens option names longer than the GA parameter value limit without collisions', async () => {
    const optionName = 'a'.repeat(101);

    await googleAnalytics.trackEvent('test command', { [optionName]: true }, { sessionId: 'session-id', shell: 'zsh' });

    const optionBody = axiosPostStub.secondCall.args[1];
    assert.strictEqual(optionBody.events[0].params.option_name.length, 100);
    assert.match(optionBody.events[0].params.option_name, /^a{91}_[0-9a-f]{8}$/);
  });

  it('sends more than 25 options as separate bounded events', async () => {
    const properties = Object.fromEntries(Array.from({ length: 27 }, (_, i) => [`option${i}`, true]));

    await googleAnalytics.trackEvent('test command', properties, { sessionId: 'session-id', shell: 'zsh' });

    assert.strictEqual(axiosPostStub.callCount, 28);
    axiosPostStub.getCalls().forEach(call => {
      assert(Object.keys(call.args[1].events[0].params).length <= 25);
    });
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