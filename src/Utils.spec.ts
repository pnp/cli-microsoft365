import * as assert from 'assert';
import Utils from './Utils';

describe('Utils', () => {

  it('isValidGuid returns true if valid guid', () => {
    
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531d6');
    assert(result);
  });

  it('isValidGuid returns false if invalid guid', () => {
    
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531dw');
    assert(result == false);
  });
});