import assert from 'assert';
import sinon from 'sinon';
import { session } from '../utils/session.js';
import { cache } from './cache.js';
import { sinonUtil } from './sinonUtil.js';

describe('utils/session', () => {
  afterEach(() => {
    sinonUtil.restore([
      cache.getValue,
      cache.setValue
    ]);
  });

  it('returns existing session ID if available', () => {
    sinon.stub(cache, 'getValue').callsFake(() => '123');
    assert.strictEqual(session.getId(1), '123');
  });

  it('returns new session ID if no ID available', () => {
    sinon.stub(cache, 'getValue').returns(undefined);
    sinon.stub(cache, 'setValue').callsFake(() => { });
    assert(session.getId(1).length > 3);
  });
});