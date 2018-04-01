import * as assert from 'assert';
import auth from './AzmgmtAuth';

describe('AzmgmtAuth', () => {
  it('uses AzMgmt service ID', () => {
    assert.equal((auth as any).serviceId(), 'AzMgmt');
  });
});