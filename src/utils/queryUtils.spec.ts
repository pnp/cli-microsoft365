import assert from 'assert';
import { queryUtils } from './queryUtils.js';

describe('utils/queryUtils', () => {
  it('createGraphQuery returns query with $select query parameter when properties argument is specified and do not contain slash', () => {
    const queryInputParameters = { properties: 'id,displayName'};
    const query = queryUtils.createGraphQuery(queryInputParameters);
    assert.strictEqual(query, '?$select=id,displayName');
  });

  it('createGraphQuery returns query with $expand query parameter when properties argument is specified and contain slash', () => {
    const queryInputParameters = { properties: 'manager/displayName,drive/id' };
    const query = queryUtils.createGraphQuery(queryInputParameters);
    assert.strictEqual(query, '?$expand=manager($select=displayName),drive($select=id)');
  });

  it('createGraphQuery returns query with $filter query parameter when filter argument is specified', () => {
    const queryInputParameters = { filter: "userType eq 'Member'" };
    const query = queryUtils.createGraphQuery(queryInputParameters);
    assert.strictEqual(query, `?$filter=userType eq 'Member'`);
  });

  it('createGraphQuery returns query with $count query parameter when count argument is set to true', () => {
    const queryInputParameters = { count: true };
    const query = queryUtils.createGraphQuery(queryInputParameters);
    assert.strictEqual(query, '?$count=true');
  });

  it('createGraphQuery returns query with $select, $expand, $filter and $count parameters', () => {
    const queryInputParameters = { properties: 'id,displayName,manager/displayName', filter: "userType eq 'Member'", count: true };
    const query = queryUtils.createGraphQuery(queryInputParameters);
    assert.strictEqual(query, `?$select=id,displayName&$expand=manager($select=displayName)&$filter=userType eq 'Member'&$count=true`);
  });
});