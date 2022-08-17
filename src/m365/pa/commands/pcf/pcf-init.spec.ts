import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import TemplateInstantiator from '../../template-instantiator';
const command: Command = require('./pcf-init');

describe(commands.PCF_INIT, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    telemetry = null;
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.readdirSync,
      TemplateInstantiator.instantiate,
      process.cwd,
      path.basename
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PCF_INIT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', () => {
    command.action(logger, { options: {} }, () => {
      assert(trackEvent.called);
    });
  });

  it('logs correct telemetry event', () => {
    command.action(logger, { options: {} }, () => {
      assert.strictEqual(telemetry.name, commands.PCF_INIT);
    });
  });

  it('supports specifying namespace', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--namespace') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying template', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--template') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation when the project directory contains relative paths', async () => {
    sinon.stub(path, 'basename').callsFake(() => 'rootPath1\\.\\..\\rootPath2');

    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the project directory equals invalid text sequences (like COM1 or LPT6)', async () => {
    sinon.stub(path, 'basename').callsFake(() => 'COM1');

    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the project directory name is emtpy', async () => {
    sinon.stub(path, 'basename').callsFake(() => '');

    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the name option isn\'t specified', async () => {
    const actual = await command.validate({ options: { namespace: 'Example.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the namespace option isn\'t specified', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the template option isn\'t specified', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported template specified', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name, namespace and Field template are specified', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name, namespace and Dataset template are specified', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Dataset' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. containing a dot)', async () => {
    const actual = await command.validate({ options: { name: 'Example1.Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. containing special character è)', async () => {
    const actual = await command.validate({ options: { name: 'Examplè1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. starting with a number)', async () => {
    const actual = await command.validate({ options: { name: '1ExampleName', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported name specified (eg. a javascript reserved word like \'innerHeight\')', async () => {
    const actual = await command.validate({ options: { name: 'innerHeight', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. first character is a dot)', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: '.Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. last character is a dot)', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace.', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. containing consecutive dots)', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1...Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. starting with a number)', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: '2Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. starting with a number after a dot)', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.2Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when unsupported namespace specified (eg. containing a javascript reserved word like \'innerHeight\')', async () => {
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.innerHeight.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the combined lengths of name and namespace exceeds 75 characters', async () => {
    const actual = await command.validate({ options: { name: 'ynnsnaclwrjxtnyzaotlrtxizfxnfyjmlzwwnetwmyxgregqzcmmwwqitoexhfftxnwbrvadhj', namespace: 'NS', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the combined lengths of name and namespace are exactly 75 characters', async () => {
    const actual = await command.validate({ options: { name: 'ynnsnaclwrjxtnyzaotlrtxizfxnfyjmlzwwnetwmyxgregqzcmmwwqitoexhfftxnwbrvadh', namespace: 'NS', template: 'Field' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the current directory doesn\'t contain any files with extension proj', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.xml', 'file3.json'] as any);
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the current directory contains files with extension proj', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.exe', 'file2.proj', 'file3.json'] as any);
    const actual = await command.validate({ options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('TemplateInstantiator.instantiate is called exactly twice', () => {
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    command.action(logger, { options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field' } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(logger, sinon.match.string, sinon.match.string, false, sinon.match.object, false).calledOnce);
      assert(templateInstantiate.withArgs(logger, sinon.match.string, sinon.match.string, true, sinon.match.object, false).calledOnce);
    });
  });

  it('TemplateInstantiator.instantiate is called exactly twice (verbose)', () => {
    const templateInstantiate = sinon.stub(TemplateInstantiator, 'instantiate').callsFake(() => { });

    command.action(logger, { options: { name: 'Example1Name', namespace: 'Example1.Namespace', template: 'Field', verbose: true } }, () => {
      assert(templateInstantiate.calledTwice);
      assert(templateInstantiate.withArgs(logger, sinon.match.string, sinon.match.string, false, sinon.match.object, true).calledOnce);
      assert(templateInstantiate.withArgs(logger, sinon.match.string, sinon.match.string, true, sinon.match.object, true).calledOnce);
    });
  });

  it('supports verbose mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});