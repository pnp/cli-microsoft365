import * as assert from 'assert';
import { fail } from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as sinon from 'sinon';
import { SinonSandbox } from 'sinon';
import { Cli, Logger } from './cli';
import Command from './Command';
import { sinonUtil } from './utils';

class SimpleCommand extends Command {
  public get name(): string {
    return 'cli mock';
  }
  public get description(): string {
    return 'Mock command';
  }
  public commandAction(logger: Logger, args: any, cb: () => void): void {
    cb();
  }
}

class CommandWithOptions extends Command {
  public get name(): string {
    return 'cli mock2';
  }
  public get description(): string {
    return 'Mock command 2';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '-l, --longOption <longOption>'
      }
    );
  }
  public commandAction(logger: Logger, args: any, cb: () => void): void {
    cb();
  }
}

class CommandWithAlias extends Command {
  public get name(): string {
    return 'cli mock';
  }
  public get description(): string {
    return 'Mock command';
  }
  public alias(): string[] | undefined {
    return ['cli alias'];
  }
  public commandAction(logger: Logger, args: any, cb: () => void): void {
    cb();
  }
}

describe('autocomplete', () => {
  let autocomplete: any;
  let sandbox: SinonSandbox;
  const commandInfo = {
    "help": {
      "--help": {}
    },
    "aad": {
    },
    "spo": {
      "app": {
      },
      "cdn": {
      },
      "connect": {
      },
      "customaction": {
      },
      "disconnect": {
      },
      "externaluser": {
      },
      "serviceprincipal": {
      },
      "sp": {
      },
      "site": {
      },
      "sitescript": {
      },
      "status": {
        "-o": [
          "json",
          "text"
        ],
        "--output": [
          "json",
          "text"
        ],
        "--verbose": {},
        "--debug": {},
        "--help": {}
      },
      "storageentity": {
      }
    }
  };
  let cli: Cli;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    autocomplete = require('./autocomplete').autocomplete;
  });

  afterEach(() => {
    (cli as any).commands = [];
  });

  after(() => {
    sinonUtil.restore([
      sandbox,
      fs.existsSync,
      fs.writeFileSync
    ]);
  });

  it('writes sh completion to disk', () => {
    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    (cli as any).loadCommand(new SimpleCommand());
    autocomplete.generateShCompletion();
    assert(writeFileSyncStub.calledWith(path.join(__dirname, `..${path.sep}commands.json`), JSON.stringify({
      cli: {
        mock: {
          "-o": ["csv", "json", "text"],
          "--query": {},
          "--output": ["csv", "json", "text"],
          "--verbose": {},
          "--debug": {},
          "--help": {},
          "-h": {}
        }
      }
    })));
  });

  it('registers sh completion using omelette', () => {
    const sandbox = sinon.createSandbox();
    const fakeOmelette = {
      setupShellInitFile: () => { }
    };
    const setupSpy = sinon.spy(fakeOmelette, 'setupShellInitFile');
    sandbox.stub(autocomplete, 'omelette').value(fakeOmelette);
    autocomplete.setupShCompletion();
    try {
      assert(setupSpy.called);
    }
    catch {
      sinonUtil.restore([
        setupSpy,
        autocomplete.omelette,
        sandbox
      ]);
    }
  });

  it('builds clink completion', () => {
    (cli as any).loadCommand(new SimpleCommand());
    const clink: string = autocomplete.getClinkCompletion();

    assert.strictEqual(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"cli"..parser({"mock"..parser({},"--debug", "--help", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-o"..parser({"csv","json","text"}))})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes long options in clink completion', () => {
    (cli as any).loadCommand(new CommandWithOptions());
    const clink: string = autocomplete.getClinkCompletion();

    assert.strictEqual(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"cli"..parser({"mock2"..parser({},"--debug", "--help", "--longOption", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-l", "-o"..parser({"csv","json","text"}))})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes short options in clink completion', () => {
    (cli as any).loadCommand(new CommandWithOptions());
    const clink: string = autocomplete.getClinkCompletion();

    assert.strictEqual(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"cli"..parser({"mock2"..parser({},"--debug", "--help", "--longOption", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-l", "-o"..parser({"csv","json","text"}))})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes autocomplete for options in clink completion', () => {
    (cli as any).loadCommand(new CommandWithOptions());
    const clink: string = autocomplete.getClinkCompletion();

    assert.strictEqual(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"cli"..parser({"mock2"..parser({},"--debug", "--help", "--longOption", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-l", "-o"..parser({"csv","json","text"}))})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes command alias in clink completion', () => {
    (cli as any).loadCommand(new CommandWithAlias());
    const clink: string = autocomplete.getClinkCompletion();

    assert.strictEqual(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"cli"..parser({"alias"..parser({},"--debug", "--help", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-o"..parser({"csv","json","text"})),"mock"..parser({},"--debug", "--help", "--output"..parser({"csv","json","text"}), "--query", "--verbose", "-h", "-o"..parser({"csv","json","text"}))})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('loads generated commands info from the file system', () => {
    sinonUtil.restore(fs.existsSync);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({}));
    (autocomplete as any).init();
    try {
      assert(readFileSyncStub.calledWith(path.join(__dirname, `..${path.sep}commands.json`), 'utf-8'));
    }
    catch (e: any) {
      fail(e);
    }
    finally {
      sinonUtil.restore([
        fs.existsSync,
        fs.readFileSync,
        readFileSyncStub
      ]);
    }
  });

  it('doesnt fail when the commands file is empty', () => {
    sinonUtil.restore(fs.existsSync);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake(() => '');
    (autocomplete as any).init();
    try {
      assert.strictEqual(JSON.stringify((autocomplete as any).commands), JSON.stringify({}));
    }
    catch (e: any) {
      fail(e);
    }
    finally {
      sinonUtil.restore([
        fs.existsSync,
        fs.readFileSync,
        readFileSyncStub
      ]);
    }
  });

  it('correctly lists available services when completing first fragment and it\'s empty', () => {
    const evtData = {
      before: "m365",
      fragment: 1,
      line: "m365 ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['help', 'aad', 'spo']));
  });

  it('correctly returns list of spo commands when first fragment is spo', () => {
    const evtData = {
      before: "spo",
      fragment: 2,
      line: "m365 spo ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['app',
      'cdn',
      'connect',
      'customaction',
      'disconnect',
      'externaluser',
      'serviceprincipal',
      'sp',
      'site',
      'sitescript',
      'status',
      'storageentity']));
  });

  it('suggests command options when line matches a command', () => {
    const evtData = {
      before: "status",
      fragment: 3,
      line: "m365 spo status ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['-o', '--output', '--verbose', '--debug', '--help']));
  });

  it('suggests option\'s values when it has autocomplete', () => {
    const evtData = {
      before: "--output",
      fragment: 4,
      line: "m365 spo status --output ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['json', 'text']));
  });

  it('suggests other available options after specifying option\'s value', () => {
    const evtData = {
      before: "json",
      fragment: 5,
      line: "m365 spo status --output json ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['-o', '--verbose', '--debug', '--help']));
  });

  it('suggests other available options if the option is a switch', () => {
    const evtData = {
      before: "--debug",
      fragment: 6,
      line: "m365 spo status --output json --debug ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith(['-o', '--verbose', '--help']));
  });

  it('doesn\'t return suggestions when the input doesn\'t match any command (completing fragment)', () => {
    const evtData = {
      before: "def",
      fragment: 2,
      line: "m365 abc def",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith([]));
  });

  it('doesn\'t return suggestions when the input doesn\'t match any command (new fragment)', () => {
    const evtData = {
      before: "def",
      fragment: 3,
      line: "m365 abc def ",
      reply: (_data: any | string[]) => { }
    };
    const replies: any[] = [];
    const replyStub = sinon.stub(evtData, 'reply').callsFake((r) => {
      replies.push(r);
    });
    autocomplete.commands = commandInfo;
    autocomplete.handleAutocomplete(undefined, evtData);
    assert(replyStub.calledWith([]));
  });
});
