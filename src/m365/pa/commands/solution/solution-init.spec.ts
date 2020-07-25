import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./solution-init');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import TemplateInstantiator from '../../template-instantiator';

describe(commands.SOLUTION_INIT, () => {
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      path.basename,
      fs.readdirSync,
      fs.existsSync,
      TemplateInstantiator.instantiate
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_INIT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', () => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      assert(trackEvent.called);
    });
  });

  it('logs correct telemetry event', () => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      assert.strictEqual(telemetry.name, commands.SOLUTION_INIT);
    });
  });

  it('supports specifying publisher name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publisherName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying publisher prefix', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publisherPrefix') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('passes validation when valid publisherName and publisherPrefix are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when the project directory contains relative paths', () => {
    sinon.stub(path, 'basename').callsFake(() => 'rootPath1\\.\\..\\rootPath2');

    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the project directory equals invalid text sequences (like COM1 or LPT6)', () => {
    sinon.stub(path, 'basename').callsFake(() => 'COM1');

    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails when the project directory name is emtpy', () => {
    sinon.stub(path, 'basename').callsFake(() => '');

    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the publisherName option isn\'t specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the publisherPrefix option isn\'t specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix is less than 2', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'p' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix is more than 8', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'verylongprefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix starts with a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: '1prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix starts with an underscore', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: '_prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix starts with \'mscrm\'', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'mscrmpr' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherPrefix contains a special character', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePublisher', publisherPrefix: 'préfix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherName starts with a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: '1ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the length of publisherName contains a special character', () => {
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: 'ExamplePùblisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the current directory doesn\'t contain any files with extension proj', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.xml', 'file3.json'] as any);
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when the current directory contains files with extension proj', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.cdsproj', 'file3.json'] as any);
    const actual = (command.validate() as CommandValidate)({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } });
    assert.notStrictEqual(actual, true);
  });

  it('TemplateInstantiator.instantiate is called exactly twice in an empty directory', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledTwice);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly twice in an empty directory (verbose)', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix', verbose: true } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, true).calledTwice);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly twice in an empty directory, using the standard publisherPrefix \'new\'', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'new' } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledTwice);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly twice when the CDS Assets Directory \'Other\' already exists in the current directory, but doesn\'t contain a Solution.xml file', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((pathToCheck) => {
      if(path.basename(pathToCheck.toString()).toLowerCase() === 'other') {
        return true;
      }
      else if (path.basename(pathToCheck.toString()).toLowerCase() === 'solution.xml') {
        return false;
      }
      else {
        return originalExistsSync(pathToCheck);
      }    
    });
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledTwice);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly once when the CDS Assets Directory \'Other\' already exists in the current directory and contains a Solution.xml file', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((pathToCheck) => {
      if(path.basename(pathToCheck.toString()).toLowerCase() === 'other') {
        return true;
      }
      else if (path.basename(pathToCheck.toString()).toLowerCase() === 'solution.xml') {
        return true;
      }
      else {
        return originalExistsSync(pathToCheck);
      }
    });
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { publisherName: '_ExamplePublisher', publisherPrefix: 'prefix' } }, () => {
      assert(templateInstantiate.calledOnce);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledOnce);
    });
  });

  it('supports verbose mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});