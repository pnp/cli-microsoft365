import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
const vorpal: Vorpal = require('./vorpal-init');
import Table = require('easy-table');
import { CommandError } from './Command';
import * as os from 'os';

describe('Utils', () => {

  it('isValidISODate returns true if value is in ISO Date format with - seperator', () => {
    const result = Utils.isValidISODate("2019-03-22");
    assert.equal(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with . seperator', () => {
    const result = Utils.isValidISODate("2019.03.22");
    assert.equal(result, true);
  });

  it('isValidISODate returns true if value is in ISO Date format with / seperator', () => {
    const result = Utils.isValidISODate("2019/03/22");
    assert.equal(result, true);
  });

  it('isValidISODate returns false if value is blank', () => {
    const result = Utils.isValidISODate("");
    assert.equal(result, false);
  });

  it('isValidISODate returns false if value is not in ISO Date format', () => {
    const result = Utils.isValidISODate("22-03-2019");
    assert.equal(result, false);
  });

  it('isValidISODate returns false if alpha characters are passed', () => {
    const result = Utils.isValidISODate("sharing is caring");
    assert.equal(result, false);
  });

  it('isValidGuid returns true if valid guid', () => {
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531d6');
    assert.equal(result, true);
  });

  it('isValidGuid returns false if invalid guid', () => {
    const result = Utils.isValidGuid('b2307a39-e878-458b-bc90-03bc578531dw');
    assert(result == false);
  });

  it('isValidBoolean returns true if valid boolean', () => {
    const result = Utils.isValidBoolean('true');
    assert.equal(result, true);
  });

  it('isValidBoolean returns false if invalid boolean', () => {
    const result = Utils.isValidBoolean('foo');
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

  it('doesn\'t fail when undefined object is passed to the log', () => {
    const actual = Utils.logOutput(undefined);
    assert.equal(actual, undefined);
  });

  it('returns the same object if non-array is passed to the log', () => {
    const s = 'foo';
    const actual = Utils.logOutput(s);
    assert.equal(actual, s);
  });

  it('doesn\'t fail when an array with undefined object is passed to the log', () => {
    const actual = Utils.logOutput([undefined]);
    assert.equal(actual, undefined);
  });

  it('formats output as JSON when JSON output requested', (done) => {
    const sandbox = sinon.createSandbox();
    if (!vorpal._command) {
      (vorpal as any)._command = undefined;
    }
    sandbox.stub(vorpal, '_command').value({
      args: {
        options: {
          output: 'json'
        }
      }
    });
    const o = { lorem: 'ipsum' };
    const actual = Utils.logOutput([o]);
    try {
      assert.equal(actual, JSON.stringify(o));
      done();
    }
    catch (e) {
      done(e);
    }
    finally {
      sandbox.restore();
    }
  });

  it('formats simple output as text', (done) => {
    const o = false;
    const actual = Utils.logOutput([o]);
    try {
      assert.equal(actual, `${o}`);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('formats date output as text', () => {
    const d = new Date();
    const actual = Utils.logOutput([d]);
    assert.equal(actual, d.toString());
  });

  it('formats object output as transposed table', (done) => {
    const o = { prop1: 'value1', prop2: 'value2' };
    const actual = Utils.logOutput([o]);
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    const expected = t.printTransposed({
      separator: ': '
    });
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('formats object output as transposed table', (done) => {
    const o = { prop1: 'value1 ', prop12: 'value12' };
    const actual = Utils.logOutput([o]);
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop12', 'value12');
    t.newRow();
    const expected = t.printTransposed({
      separator: ': '
    });
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('formats array values as JSON', (done) => {
    const o = { prop1: ['value1', 'value2'] };
    const actual = Utils.logOutput([o]);
    const expected = 'prop1: ["value1","value2"]' + '\n';
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('formats array output as table', (done) => {
    const o = [
      { prop1: 'value1', prop2: 'value2' },
      { prop1: 'value3', prop2: 'value4' }
    ];
    const actual = Utils.logOutput([o]);
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    t.cell('prop1', 'value3');
    t.cell('prop2', 'value4');
    t.newRow();
    const expected = t.toString();
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('formats command error as error message', (done) => {
    const o = new CommandError('An error has occurred');
    const actual = Utils.logOutput([o]);
    const expected = vorpal.chalk.red('Error: An error has occurred');
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('sets array type to the first non-undefined value', (done) => {
    const o = [undefined, 'lorem', 'ipsum'];
    const actual = Utils.logOutput([o]);
    const expected = `${os.EOL}lorem${os.EOL}ipsum`;
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('skips primitives mixed with objects when rendering a table', (done) => {
    const o = [
      { prop1: 'value1', prop2: 'value2' },
      'lorem',
      { prop1: 'value3', prop2: 'value4' }
    ];
    const actual = Utils.logOutput([o]);
    const t = new Table();
    t.cell('prop1', 'value1');
    t.cell('prop2', 'value2');
    t.newRow();
    t.cell('prop1', 'value3');
    t.cell('prop2', 'value4');
    t.newRow();
    const expected = t.toString();
    try {
      assert.equal(actual, expected);
      done();
    }
    catch (e) {
      done(e);
    }
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1');
    assert.equal(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/team1/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/');
    assert.equal(actual, '/sites/team1');
  });

  it('should get server relative path when https://contoso.sharepoint.com/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/');
    assert.equal(actual, '/');
  });

  it('should get server relative path when domain only', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com');
    assert.equal(actual, '/');
  });

  it('should get server relative path when /sites/team1 relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/sites/team1');
    assert.equal(actual, '/sites/team1');
  });

  it('should get server relative path when /sites/team1/ relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/sites/team1/');
    assert.equal(actual, '/sites/team1');
  });

  it('should get server relative path when / relative path passed as param', () => {
    const actual = Utils.getServerRelativePath('/');
    assert.equal(actual, '/');
  });

  it('should get server relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when domain only and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/team1', '/Shared Documents');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/team1', '/Shared Documents/');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/team1/', '/Shared Documents/');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/team1/', 'Shared Documents');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/team1', 'Shared Documents');
    assert.equal(actual, '/sites/team1/Shared Documents');
  });

  it('should get server relative path when / and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when / and /Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/', '/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when "" and ""', () => {
    const actual = Utils.getServerRelativePath('', '');
    assert.equal(actual, '/');
  });

  it('should get server relative path when / and /', () => {
    const actual = Utils.getServerRelativePath('/', '/');
    assert.equal(actual, '/');
  });

  it('should get server relative path when "" and /', () => {
    const actual = Utils.getServerRelativePath('', '/');
    assert.equal(actual, '/');
  });

  it('should get server relative path when "" and Shared Documents', () => {
    const actual = Utils.getServerRelativePath('', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getServerRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/sites/Site1/Shared Documents');
  });

  it('should get server relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.equal(actual, '/sites/site1/Shared Documents');
  });

  it('should get server relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = Utils.getServerRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.equal(actual, '/sites/site1/Shared Documents/MyFolder');
  });

  it('should get web relative path when / relative path passed as param', () => {
    const actual = Utils.getWebRelativePath('/', '/');
    assert.equal(actual, '/');
  });

  it('should get web relative path for https://contoso.sharepoint.com/sites/team1 and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/team1', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/team1/', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/ and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when domain only and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/team1', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1 and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/team1', '/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/team1/ and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/team1/', '/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/team1/', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/team1/ and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/team1', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when / and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/', '/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when / and /Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/', '/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when "" and ""', () => {
    const actual = Utils.getWebRelativePath('', '');
    assert.equal(actual, '/');
  });

  it('should get web relative path when / and /', () => {
    const actual = Utils.getWebRelativePath('/', '/');
    assert.equal(actual, '/');
  });

  it('should get web relative path when "" and /', () => {
    const actual = Utils.getWebRelativePath('', '/');
    assert.equal(actual, '/');
  });

  it('should get web relative path when "" and Shared Documents', () => {
    const actual = Utils.getWebRelativePath('', 'Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1/', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when https://contoso.sharepoint.com/sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('https://contoso.sharepoint.com/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and /sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1/', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1/ and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when /sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('/sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sites/site1 and sites/site1/Shared Documents/', () => {
    const actual = Utils.getWebRelativePath('sites/site1', 'sites/site1/Shared Documents/');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in web url e.g. sites/Site1 and /sites/site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/Site1', '/sites/site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when uppercase in folder url e.g. sites/site1 and /sites/Site1/Shared Documents', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents');
    assert.equal(actual, '/Shared Documents');
  });

  it('should get web relative path when sub folder present url e.g. sites/site1 and /sites/Site1/Shared Documents/MyFolder', () => {
    const actual = Utils.getWebRelativePath('sites/site1', '/sites/Site1/Shared Documents/MyFolder');
    assert.equal(actual, '/Shared Documents/MyFolder');
  });

  it('should get absolute URL of a folder using webUrl and the folder server relative url', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', '/sites/team1/Shared Documents/MyFolder');
    assert.equal(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the server relative url starting by / or not while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1', 'sites/team1/Shared Documents/MyFolder');
    assert.equal(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should handle the presence of an ending / on the web url while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', 'sites/team1/Shared Documents/MyFolder');
    assert.equal(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('should properly concatenate URL parts even with ending and starting / to each while getting absolute URL of a folder', () => {
    const actual = Utils.getAbsoluteUrl('https://contoso.sharepoint.com/sites/team1/', '/sites/team1/Shared Documents/MyFolder');
    assert.equal(actual, 'https://contoso.sharepoint.com/sites/team1/Shared Documents/MyFolder');
  });

  it('shows app display name as connected-as for app-only auth', () => {
    const jwt = JSON.stringify({
      app_displayname: 'Office 365 CLI Contoso'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;
    const actual = Utils.getUserNameFromAccessToken(accessToken);
    assert.equal(actual, 'Office 365 CLI Contoso');
  });
});