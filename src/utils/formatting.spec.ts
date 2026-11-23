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

  it('casts numeric-looking values to numbers by default when using parseCsvToJson', () => {
    const result = formatting.parseCsvToJson('Title,Count\nA,5');
    assert.deepStrictEqual(result, [{ Title: 'A', Count: 5 }]);
  });

  it('keeps values as text when cast is disabled when using parseCsvToJson', () => {
    const result = formatting.parseCsvToJson('Title,Zip,Phone,Code,Exp\nA,"00123",+31612345678,12345678901234567890,1e3', { cast: false });
    assert.deepStrictEqual(result, [{ Title: 'A', Zip: '00123', Phone: '+31612345678', Code: '12345678901234567890', Exp: '1e3' }]);
  });

  it('strips the UTF-8 BOM from the first header when using parseCsvToJson', () => {
    const result = formatting.parseCsvToJson('\uFEFFID,Title\n1,A');
    assert.deepStrictEqual(Object.keys(result[0]), ['ID', 'Title']);
  });

  it('uses the specified quote character and delimiter when using parseCsvToJson', () => {
    const result = formatting.parseCsvToJson('Title;Note\n|A;B|;C', { quoteChar: '|', delimiter: ';' });
    assert.deepStrictEqual(result, [{ Title: 'A;B', Note: 'C' }]);
  });
});
