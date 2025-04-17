import assert from 'assert';
import { formatting } from './formatting.js';

describe('utils/formatting', () => {
  it('correctly returns GUID value from a CSOM GUID string when using extractCsomGuid', () => {
    const result = formatting.extractCsomGuid('/Guid(5c51a9d1-0f07-4e61-879e-0a286568c232)/');
    assert.strictEqual(result, '5c51a9d1-0f07-4e61-879e-0a286568c232');
  });

  it('correctly returns GUID value from a CSOM GUID string in capitals when using extractCsomGuid', () => {
    const result = formatting.extractCsomGuid('/GUID(5C51A9D1-0F07-4E61-879E-0A286568C232)/');
    assert.strictEqual(result, '5C51A9D1-0F07-4E61-879E-0A286568C232');
  });

  it('correctly returns default value from a when using invalid GUID string for extractCsomGuid', () => {
    const result = formatting.extractCsomGuid('invalid');
    assert.strictEqual(result, 'invalid');
  });
});