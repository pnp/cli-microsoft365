import assert from 'assert';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './project-doctor.js';
import { FindingToReport } from './report-model/index.js';

describe(commands.PROJECT_DOCTOR, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let trackEvent: any;
  let telemetryCommandName: any;
  const validProjectPath = 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-react';
  const invalidProjectPath = 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-react-invalidconfig';

  before(() => {
    trackEvent = sinon.stub(telemetry, 'trackEvent').callsFake((commandName) => {
      telemetryCommandName = commandName;
      return Promise.resolve();
    });
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    telemetryCommandName = null;
    (command as any).allFindings = [];
    (command as any).packageManager = 'npm';
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      fs.statSync,
      fs.writeFileSync,
      fs.mkdirSync,
      fsUtil.readdirR
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_DOCTOR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: {} });
    assert(trackEvent.called);
  });

  it('logs correct telemetry event', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: {} });
    assert.strictEqual(telemetryCommandName, commands.PROJECT_DOCTOR);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(null);

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(`Couldn't find project root folder`, 1));
  });

  it('shows error if the project version couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').returns(undefined);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Unable to determine the version of the current SharePoint Framework project`, 3));
  });

  it('shows error if the project version is not supported', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').returns('0.0.1');

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`CLI for Microsoft 365 doesn't support validating projects built using SharePoint Framework v0.0.1`, 4));
  });

  it('shows error when loading doctor rules failed', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').returns('0');

    (command as any).supportedVersions.splice(1, 0, '0');

    await assert.rejects(command.action(logger, { options: {} } as any), (err) => {
      (command as any).supportedVersions.splice(1, 1);
      const message = (err as any).message;
      return message.indexOf('Cannot find module') > -1 &&
        message.indexOf(`${path.sep}project-doctor${path.sep}doctor-0.js'`) > -1;
    });
  });

  it('returns markdown report with output format md', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    sinon.stub(cli, 'log').callsFake(msg => log.push(msg));

    try {
      await cli.executeCommand(command, { options: { output: 'md' } } as any);
      assert(log[0].indexOf('## Findings') > -1);
    }
    finally {
      sinonUtil.restore(cli.log);
    }
  });

  it('overrides base md formatting', async () => {
    const expected = [
      {
        'prop1': 'value1'
      },
      {
        'prop2': 'value2'
      }
    ];
    const actual = command.getMdOutput(expected, command, { options: { output: 'md' } } as any);
    assert.deepStrictEqual(actual, expected);
  });

  it('returns text report with output format text', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: { output: 'text' } } as any);
    assert(log[0].indexOf('-----------------------') > -1);
  });

  it('returns json report with output format default', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: {} } as any);
    assert(Array.isArray(log[0]));
  });

  it('writes CodeTour doctor report to .tours folder when in tour output mode. Creates the folder when it does not exist', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return false;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').resolves('');

    await command.action(logger, { options: { output: 'tour' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours/validation.tour')), 'Tour file not created');
    assert(mkDirSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours')), '.tours folder not created');
  });

  it('writes CodeTour upgrade report to .tours folder when in tour output mode. Does not create the folder when it already exists', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return true;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').resolves('');

    await command.action(logger, { options: { output: 'tour' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours/validation.tour')), 'Tour file not created');
    assert(mkDirSyncStub.notCalled, '.tours folder created');
  });

  it('e2e: shows correct number of findings for a valid 1.0.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 7);
  });

  it('e2e: shows correct number of findings for a valid 1.0.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 7);
  });

  it('e2e: shows correct number of findings for a valid 1.0.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 7);
  });

  it('e2e: shows correct number of findings for a valid 1.1.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for a valid 1.1.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for a valid 1.1.3 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for a valid 1.2.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for a valid 1.3.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for a valid 1.3.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for a valid 1.3.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for a valid 1.3.4 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for a valid 1.4.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for a valid 1.4.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for a valid 1.5.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.5.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.6.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.7.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.8.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.8.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.8.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.9.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.10.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for a valid 1.11.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.12.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.12.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.13.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.13.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.14.0 project (json)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), validProjectPath));

    await command.action(logger, { options: { output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct message a valid 1.14.0 project (text)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), validProjectPath));

    await command.action(logger, { options: { output: 'text' } } as any);
    assert.strictEqual(log[0], '✅ CLI for Microsoft 365 has found no issues in your project');
  });

  it('e2e: shows correct message for a valid 1.14.0 project (md)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), validProjectPath));

    await command.action(logger, { options: { output: 'md' } } as any);
    assert(log[0].indexOf('✅ CLI for Microsoft 365 has found no issues in your project') > -1);
  });

  it('e2e: shows yarn commands for yarn package manager', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: { output: 'json', packageManager: 'yarn' } } as any);
    const findings: FindingToReport[] = log.pop();
    assert.strictEqual(findings[0].resolution.indexOf('yarn '), 0);
  });

  it('e2e: shows yarn commands for pnpm package manager', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: { output: 'json', packageManager: 'pnpm' } } as any);
    const findings: FindingToReport[] = log.pop();
    assert.strictEqual(findings[0].resolution.indexOf('pnpm '), 0);
  });

  it('e2e: shows correct number of findings for an invalid 1.14.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: { output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 28);
  });

  it('e2e: shows correct number of findings for an invalid 1.14.0 project (debug)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), invalidProjectPath));

    await command.action(logger, { options: { output: 'json', debug: true } } as any);
    const findings: FindingToReport[] = log.pop();
    assert.strictEqual(findings.length, 28);
  });

  it('e2e: shows correct number of findings for a valid 1.15.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.15.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.16.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.16.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1161-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.17.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1170-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.17.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1171-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.17.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1172-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 2);
  });

  it('e2e: shows correct number of findings for a valid 1.17.3 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1173-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.17.4 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1174-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.18.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1180-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.18.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1181-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.18.2 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1182-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.19.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1190-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.20.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1200-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.21.0 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1210-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('e2e: shows correct number of findings for a valid 1.22.0-beta.1 project', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1220-beta.1-webpart-react'));

    await command.action(logger, { options: {} } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 0);
  });

  it('passes validation when package manager not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when npm package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'npm' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when pnpm package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'pnpm' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when yarn package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'yarn' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when json output specified', async () => {
    assert.strictEqual(await command.validate({ options: { output: 'json' } }, cli.getCommandInfo(command)), true);
  });

  it('passes validation when text output specified', async () => {
    assert.strictEqual(await command.validate({ options: { output: 'text' } }, cli.getCommandInfo(command)), true);
  });

  it('passes validation when md output specified', async () => {
    assert.strictEqual(await command.validate({ options: { output: 'md' } }, cli.getCommandInfo(command)), true);
  });

  it('passes validation when tour output specified', async () => {
    assert.strictEqual(await command.validate({ options: { output: 'tour' } }, cli.getCommandInfo(command)), true);
  });

  it('fails validation when csv output specified', async () => {
    assert.notStrictEqual(await command.validate({ options: { output: 'csv' } }, cli.getCommandInfo(command)), true);
  });
});
