import assert from 'assert';
import Axios from 'axios';
import Configstore from 'configstore';
import sinon from 'sinon';
import { googleAnalytics } from './googleAnalytics.js';

describe('googleAnalytics', () => {
  let axiosPostStub: sinon.SinonStub;

  beforeEach(() => {
    sinon.stub(Configstore.prototype, 'get').returns('client-id');
    axiosPostStub = sinon.stub(Axios, 'post').resolves({ status: 204 });
  });

  afterEach(() => {
    sinon.restore();
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
    assert.strictEqual(body.events[0].params.debug, false);
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

  it('throws when Google Analytics rejects the request', async () => {
    axiosPostStub.resolves({ status: 500 });

    await assert.rejects(
      googleAnalytics.trackEvent('test command', {}, { sessionId: 'session-id', shell: 'zsh' }),
      /Google Analytics returned 500/
    );
  });
});