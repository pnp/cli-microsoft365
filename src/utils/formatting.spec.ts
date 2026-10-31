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

  it('correctly returns ISO date from a CSOM date string when using extractCsomDate', () => {
    const result = formatting.extractCsomDate('/Date(2026,8,21,12,24,45,997)/');
    assert.strictEqual(result, '2026-09-21T12:24:45.997Z');
  });

  it('correctly returns ISO date from an escaped CSOM date string with surrounding whitespace when using extractCsomDate', () => {
    const result = formatting.extractCsomDate('\\/Date(2026,8,21,12,24,45,997)\\/ ');
    assert.strictEqual(result, '2026-09-21T12:24:45.997Z');
  });

  it('correctly returns ISO date from a CSOM date string without time parts when using extractCsomDate', () => {
    const result = formatting.extractCsomDate('/Date(2026,0,1)/');
    assert.strictEqual(result, '2026-01-01T00:00:00.000Z');
  });

  it('correctly returns original value when using invalid date string for extractCsomDate', () => {
    const result = formatting.extractCsomDate('invalid');
    assert.strictEqual(result, 'invalid');
  });
});