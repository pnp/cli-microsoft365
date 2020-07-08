import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./pcf-init');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import TemplateInstantiator from '../../template-instantiator';

describe(commands.PCF_INIT, () => {
  let vorpal: Vorpal;
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
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
      fs.readdirSync,
      TemplateInstantiator.instantiate,
      process.cwd,
      path.basename
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PCF_INIT), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
      assert.equal(telemetry.name, commands.PCF_INIT);
    });
  });

  it('supports specifying namespace', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--namespace') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying template', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--template') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation when the project directory contains relative paths', () => {
    sinon.stub(path, 'basename').callsFake(() => 'rootPath1\\.\\..\\rootPath2');

    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the project directory equals invalid text sequences (like COM1 or LPT6)', () => {
    sinon.stub(path, 'basename').callsFake(() => 'COM1');

    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails when the project directory name is emtpy', () => {
    sinon.stub(path, 'basename').callsFake(() => '');

    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the name option isn\'t specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { namespace: 'Example.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the namespace option isn\'t specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the template option isn\'t specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported template specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when name, namespace and Field template are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.equal(actual, true);
  });

  it('passes validation when name, namespace and Dataset template are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Dataset' } });
    assert.equal(actual, true);
  });

  it('fails validation when unsupported name specified (eg. containing a dot)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1.Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. containing special character è)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Examplè1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. starting with a number)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: '1ExampleName', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. a javascript reserved word like \'innerHeight\')', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'innerHeight', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. first character is a dot)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: '.Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. last character is a dot)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace.', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. containing consecutive dots)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1...Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. starting with a number)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: '2Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. starting with a number after a dot)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.2Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. containing a javascript reserved word like \'innerHeight\')', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.innerHeight.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the combined lengths of name and namespace exceeds 75 characters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'ynnsnaclwrjxtnyzaotlrtxizfxnfyjmlzwwnetwmyxgregqzcmmwwqitoexhfftxnwbrvadhj', namespace: 'NS', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the combined lengths of name and namespace are exactly 75 characters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'ynnsnaclwrjxtnyzaotlrtxizfxnfyjmlzwwnetwmyxgregqzcmmwwqitoexhfftxnwbrvadh', namespace: 'NS', template: 'Field' } });
    assert.equal(actual, true);
  });

  it('passes validation when the current directory doesn\'t contain any files with extension proj', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.xml', 'file3.json'] as any);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.equal(actual, true);
  });

  it('fails validation when the current directory contains files with extension proj', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.proj', 'file3.json'] as any);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } });
    assert.notEqual(actual, true);
  });

  it('TemplateInstantiator.instantiate is called exactly twice', () => {
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledOnce);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, true, sinon.match.object, false).calledOnce);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly twice (verbose)', () => {
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    cmdInstance.action({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field', verbose: true } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, false, sinon.match.object, true).calledOnce);
      assert(templateInstantiate.withArgs(cmdInstance, sinon.match.string, sinon.match.string, true, sinon.match.object, true).calledOnce);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.PCF_INIT));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});