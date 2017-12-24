import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
const vorpal: Vorpal = require('./vorpal-init');
import Table = require('easy-table');
import { CommandError } from './Command';
import * as os from 'os';

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
    const expected = 'prop1: ["value1","value2"]' + os.EOL;
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
});