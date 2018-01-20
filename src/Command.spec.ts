import * as sinon from 'sinon';
import * as assert from 'assert';
import Command, {
  CommandValidate,
  CommandCancel,
  CommandOption,
  CommandTypes
} from './Command';
import Utils from './Utils';
const vorpal: Vorpal = require('./vorpal-init');

class MockCommand1 extends Command {
  public get name(): string {
    return 'Mock command';
  }

  public get description(): string {
    return 'Mock command description';
  }

  public alias(): string[] | undefined {
    return ['mc1'];
  }

  public autocomplete(): string[] | undefined {
    const autocomplete = ['param1', 'param2'];

    const parentAutocomplete: string[] | undefined = super.autocomplete();
    if (parentAutocomplete) {
      return autocomplete.concat(parentAutocomplete);
    }
    else {
      return autocomplete;
    }
  }

  public allowUnknownOptions(): boolean {
    return true;
  }

  public commandAction(): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }

  public validate(): CommandValidate | undefined {
    return () => {
      return true;
    };
  }

  public cancel(): CommandCancel | undefined {
    return () => { };
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['']
    };
  }

  public options(): CommandOption[] {
    return [{
      option: '--debug',
      description: 'Runs command with debug logging'
    }];
  }
}

class MockCommand2 extends Command {
  public get name(): string {
    return 'Mock command 2 [opt]';
  }

  public get description(): string {
    return 'Mock command 2 description';
  }

  public commandAction(): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
    log('MockCommand2 help');
  }
}

describe('Command', () => {
  const vcmd = {
    action: () => vcmd,
    alias: () => vcmd,
    option: () => vcmd,
    validate: () => vcmd,
    cancel: () => vcmd,
    help: () => vcmd,
    types: () => vcmd,
    allowUnknownOptions: () => vcmd
  };
  let actionSpy: sinon.SinonSpy;
  let aliasSpy: sinon.SinonSpy;
  let optionSpy: sinon.SinonSpy;
  let validateSpy: sinon.SinonSpy;
  let cancelSpy: sinon.SinonSpy;
  let helpSpy: sinon.SinonSpy;
  let typesSpy: sinon.SinonSpy;

  beforeEach(() => {
    actionSpy = sinon.spy(vcmd, 'action');
    aliasSpy = sinon.spy(vcmd, 'alias');
    optionSpy = sinon.spy(vcmd, 'option');
    validateSpy = sinon.spy(vcmd, 'validate');
    cancelSpy = sinon.spy(vcmd, 'cancel');
    helpSpy = sinon.spy(vcmd, 'help');
    typesSpy = sinon.spy(vcmd, 'types');
  });

  afterEach(() => {
    Utils.restore([
      vcmd.action,
      vcmd.alias,
      vcmd.option,
      vcmd.validate,
      vcmd.cancel,
      vcmd.help,
      vcmd.types,
      vcmd.allowUnknownOptions
    ]);
  });

  it('initiates command 1 with vorpal', () => {
    const cmd = new MockCommand1();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command 2 with vorpal', () => {
    const cmd = new MockCommand2();
    const vorpalCommandStub = sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(vorpalCommandStub.calledOnce);
  });

  it('initiates command with command name', () => {
    const cmd = new MockCommand1();
    let name;
    sinon.stub(vorpal, 'command').callsFake((_name) => {
      name = _name;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.equal(name, cmd.name);
  });

  it('initiates command with command description', () => {
    const cmd = new MockCommand1();
    let description;
    sinon.stub(vorpal, 'command').callsFake((_name, _description) => {
      description = _description;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.equal(description, cmd.description);
  });

  it('initiates command with command autocomplete', () => {
    const cmd = new MockCommand1();
    let autocomplete;
    sinon.stub(vorpal, 'command').callsFake((_name, _description, _autocomplete) => {
      autocomplete = _autocomplete;
      return vcmd;
    });
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert.deepEqual(autocomplete, cmd.autocomplete());
  });

  it('configures command action', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(actionSpy.calledOnce);
  });

  it('configures options when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(optionSpy.calledOnce);
  });

  it('configures alias when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(aliasSpy.calledOnce);
  });

  it('configures validation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.calledOnce);
  });

  it('doesn\'t configure validation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(validateSpy.notCalled);
  });

  it('configures cancellation when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.calledOnce);
  });

  it('doesn\'t configure cancellation when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(cancelSpy.notCalled);
  });

  it('configures help when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(helpSpy.calledOnce);
  });

  it('configures types when available', () => {
    const cmd = new MockCommand1();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(typesSpy.calledOnce);
  });

  it('doesn\'t configure type when unavailable', () => {
    const cmd = new MockCommand2();
    sinon.stub(vorpal, 'command').callsFake(() => vcmd);
    cmd.init(vorpal);
    Utils.restore(vorpal.command);
    assert(typesSpy.notCalled);
  });

  it('returns command name without arguments', () => {
    const cmd = new MockCommand2();
    assert.equal(cmd.getCommandName(), 'Mock command 2');
  });

  it('prints help using the log argument when called from the help command', () => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(vorpal, '_command').value({
      command: 'help mock2'
    });
    const log = (msg?: string) => { };
    const logSpy = sinon.spy(log);
    const mock = new MockCommand2();
    const cmd = {
      help: mock.help()
    };
    cmd.help({}, logSpy);
    sandbox.restore();
    assert(logSpy.called);
  })
});