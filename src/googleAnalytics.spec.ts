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

  it('sends command and option usage using Google Analytics collection', async () => {
    await googleAnalytics.trackEvent('spo file add', { output: 'json', debug: false }, {
      sessionId: 'session-id',
      shell: 'zsh'
    });

    assert.strictEqual(axiosPostStub.callCount, 3);
    assert.strictEqual(axiosPostStub.firstCall.args[0], 'https://www.google-analytics.com/g/collect');

    const commandEvent = new URLSearchParams(axiosPostStub.firstCall.args[1]);
    assert.strictEqual(commandEvent.get('tid'), 'G-4BNT8MQCYT');
    assert.strictEqual(commandEvent.get('cid'), 'client-id');
    assert.strictEqual(commandEvent.get('en'), 'command_used');
    assert.strictEqual(commandEvent.get('ep.command_name'), 'spo file add');
    assert.strictEqual(commandEvent.get('ep.shell'), 'zsh');
    assert.strictEqual(commandEvent.get('ep.os'), process.platform);

    const outputEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(outputEvent.get('en'), 'command_option_used');
    assert.strictEqual(outputEvent.get('ep.option_name'), 'output');
    assert.strictEqual(outputEvent.get('ep.option_value'), 'json');

    const debugEvent = new URLSearchParams(axiosPostStub.thirdCall.args[1]);
    assert.strictEqual(debugEvent.get('ep.option_name'), 'debug');
    assert.strictEqual(debugEvent.get('ep.option_value'), 'false');

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
    const outputEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(outputEvent.get('ep.option_name'), 'output');
    assert.strictEqual(outputEvent.get('ep.option_value'), 'json');
  });

  it('converts non-primitive option properties to strings', async () => {
    await googleAnalytics.trackEvent('test command', { value: null }, { sessionId: 'session-id', shell: 'zsh' });

    const optionEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(optionEvent.get('ep.option_value'), 'null');
  });

  it('sends numeric option properties as numbers', async () => {
    await googleAnalytics.trackEvent('test command', { value: 1 }, { sessionId: 'session-id', shell: 'zsh' });

    const optionEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(optionEvent.get('epn.option_value'), '1');
  });

  it('sends option names longer than 40 characters as parameter values', async () => {
    const optionName = 'disableTrialEnvironmentCreationByNonAdminUsers';

    await googleAnalytics.trackEvent('test command', { [optionName]: true }, { sessionId: 'session-id', shell: 'zsh' });

    const optionEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(optionEvent.get('ep.option_name'), optionName);
    assert.strictEqual(optionEvent.get('ep.option_value'), 'true');
  });

  it('shortens option names longer than the GA parameter value limit without collisions', async () => {
    const optionName = 'a'.repeat(101);

    await googleAnalytics.trackEvent('test command', { [optionName]: true }, { sessionId: 'session-id', shell: 'zsh' });

    const optionEvent = new URLSearchParams(axiosPostStub.secondCall.args[1]);
    assert.strictEqual(optionEvent.get('ep.option_name')!.length, 100);
    assert.match(optionEvent.get('ep.option_name')!, /^a{91}_[0-9a-f]{8}$/);
  });

  it('sends more than 25 options as separate bounded events', async () => {
    const properties = Object.fromEntries(Array.from({ length: 27 }, (_, i) => [`option${i}`, true]));

    await googleAnalytics.trackEvent('test command', properties, { sessionId: 'session-id', shell: 'zsh' });

    assert.strictEqual(axiosPostStub.callCount, 28);
    axiosPostStub.getCalls().forEach(call => {
      const event = new URLSearchParams(call.args[1]);
      const eventParameterCount = Array.from(event.keys()).filter(key => key.startsWith('ep.') || key.startsWith('epn.')).length;
      assert(eventParameterCount <= 25);
    });
  });

  it('sets the Docker environment in telemetry', async () => {
    process.env.CLIMICROSOFT365_ENV = 'docker';
    const modulePath = `./googleAnalytics.js?docker=${Math.random()}`;
    const { googleAnalytics: dockerGoogleAnalytics } = await import(modulePath);

    assert.strictEqual(dockerGoogleAnalytics.commonProperties.env, 'docker');
  });

  it('sends errors as exceptions using Google Analytics collection', async () => {
    const message = 'a'.repeat(101);

    await googleAnalytics.trackException(new Error(message), { sessionId: 'session-id', shell: 'zsh' });

    const exceptionEvent = new URLSearchParams(axiosPostStub.firstCall.args[1]);
    assert.strictEqual(exceptionEvent.get('en'), 'exception');
    assert.strictEqual(exceptionEvent.get('ep.description'), message.substring(0, 100));
    assert.strictEqual(exceptionEvent.get('ep.shell'), 'zsh');
  });

  it('converts non-error exceptions to strings', async () => {
    await googleAnalytics.trackException('error', { sessionId: 'session-id', shell: 'zsh' });

    const exceptionEvent = new URLSearchParams(axiosPostStub.firstCall.args[1]);
    assert.strictEqual(exceptionEvent.get('ep.description'), 'error');
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