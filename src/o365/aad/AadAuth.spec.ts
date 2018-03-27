import * as assert from 'assert';
import auth from './AadAuth';

describe('AadAuth', () => {
  it('uses AAD service ID', () => {
    assert.equal((auth as any).serviceId(), 'AAD');
  });
});