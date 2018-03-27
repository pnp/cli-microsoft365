import * as assert from 'assert';
import auth from './GraphAuth';

describe('GraphAuth', () => {
  it('uses Graph service ID', () => {
    assert.equal((auth as any).serviceId(), 'Graph');
  });
});