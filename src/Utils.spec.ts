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

  it('adds User-Agent string to undefined headers', () => {
    const result = Utils.getRequestHeaders(undefined);
    assert.equal(!result['User-Agent'], false);
  });

  it('adds User-Agent string to empty headers', () => {
    const result = Utils.getRequestHeaders({});
    assert.equal(!result['User-Agent'], false);
  });

  it('adds User-Agent string to existing headers', () => {
    const result = Utils.getRequestHeaders({ accept: 'application/json' });
    assert.equal(!result['User-Agent'], false);
    assert.equal(result.accept, 'application/json');
  });

  it('doesn\'t fail when restoring stub if the passed object is undefined', () => {
    Utils.restore(undefined);
    assert(true);
  });
});