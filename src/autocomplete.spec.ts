import * as sinon from 'sinon';
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import Utils from './Utils';
import { SinonSandbox } from 'sinon';
import { fail } from 'assert';

describe('autocomplete', () => {
  let autocomplete: any;
  let sandbox: SinonSandbox;
  let commandInfo = {
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

  before(() => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    autocomplete = require('./autocomplete').autocomplete;
  });

  after(() => {
    Utils.restore([
      sandbox,
      fs.existsSync,
      fs.writeFileSync
    ]);
  });

  it('writes sh completion to disk', () => {
    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake((path, contents) => { });
    autocomplete.generateShCompletion({
      commands: [
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'spo connect',
          _hidden: false
        }
      ]
    });
    assert(writeFileSyncStub.calledWith(path.join(__dirname, `..${path.sep}commands.json`), JSON.stringify({
      spo: {
        connect: {
          '--help': {}
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
      Utils.restore([
        setupSpy,
        autocomplete.omelette,
        sandbox
      ]);
    }
  });

  it('builds clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'spo connect',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"connect"..parser({},"--help")})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('ignores the exit command in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'spo connect',
          _hidden: false
        },
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'exit',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"connect"..parser({},"--help")})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('ignores the quit command in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'spo connect',
          _hidden: false
        },
        {
          options: [],
          _args: [],
          _aliases: [],
          _name: 'quit',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"connect"..parser({},"--help")})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes long options in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [
            {
              autocomplete: null,
              long: '--appCatalogUrl',
              short: null
            }
          ],
          _args: [],
          _aliases: [],
          _name: 'spo app list',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"app"..parser({"list"..parser({},"--appCatalogUrl", "--help")})})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes short options in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [
            {
              autocomplete: [],
              long: null,
              short: "-u"
            }
          ],
          _args: [],
          _aliases: [],
          _name: 'spo app list',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"app"..parser({"list"..parser({},"--help", "-u")})})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes autocomplete for options in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [
            {
              autocomplete: ['json', 'text'],
              long: null,
              short: "-o"
            }
          ],
          _args: [],
          _aliases: [],
          _name: 'spo app list',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"app"..parser({"list"..parser({},"--help", "-o"..parser({"json","text"}))})})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('includes command alias in clink completion', () => {
    const clink: string = autocomplete.getClinkCompletion({
      commands: [
        {
          options: [],
          _args: [],
          _aliases: ['spo c'],
          _name: 'spo connect',
          _hidden: false
        }
      ]
    });

    assert.equal(clink, [
      'local parser = clink.arg.new_parser',
      'local m365_parser = parser({"spo"..parser({"c"..parser({},"--help"),"connect"..parser({},"--help")})})',
      '',
      'clink.arg.register_parser("m365", m365_parser)',
      'clink.arg.register_parser("microsoft365", m365_parser)'
    ].join(os.EOL));
  });

  it('loads generated commands info from the file system', () => {
    Utils.restore(fs.existsSync);
    sinon.stub(fs, 'existsSync').callsFake((path) => true);
    const readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => JSON.stringify({}));
    (autocomplete as any).init();
    try {
      assert(readFileSyncStub.calledWith(path.join(__dirname, `..${path.sep}commands.json`), 'utf-8'));
    }
    catch (e) {
      fail(e);
    }
    finally {
      Utils.restore([
        fs.existsSync,
        fs.readFileSync,
        readFileSyncStub
      ]);
    }
  });

  it('doesnt fail when the commands file is empty', () => {
    Utils.restore(fs.existsSync);
    sinon.stub(fs, 'existsSync').callsFake((path) => true);
    const readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => '');
    (autocomplete as any).init();
    try {
      assert.equal(JSON.stringify((autocomplete as any).commands), JSON.stringify({}));
    }
    catch (e) {
      fail(e);
    }
    finally {
      Utils.restore([
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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
      reply: (data: Object | string[]) => { }
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