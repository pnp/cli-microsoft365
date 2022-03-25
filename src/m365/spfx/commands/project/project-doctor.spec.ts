import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import { fsUtil, sinonUtil } from '../../../../utils';
import commands from '../../commands';
import { FindingToReport } from './report-model';
const command: Command = require('./project-doctor');

describe(commands.PROJECT_DOCTOR, () => {
  let log: any[];
  let logger: Logger;
  let trackEvent: any;
  let telemetry: any;
  const validProjectPath = 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-react';
  const invalidProjectPath = 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-react-invalidconfig';

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
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
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROJECT_DOCTOR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: {} }, () => {
      try {
        assert.strictEqual(telemetry.name, commands.PROJECT_DOCTOR);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error if the project path couldn\'t be determined', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error if the project version couldn\'t be determined', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => undefined);

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, 3)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error if the project version is not supported', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '0.0.1');

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`CLI for Microsoft 365 doesn't support validating projects built using SharePoint Framework v0.0.1`, 4)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when loading doctor rules failed', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '0');

    (command as any).supportedVersions.splice(1, 0, '0');

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        (command as any).supportedVersions.splice(1, 1);
        assert(JSON.stringify(err).indexOf("Cannot find module './project-doctor/doctor-0'") > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns markdown report with output format md', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'md' } } as any, () => {
      try {
        assert(log[0].indexOf('## Findings') > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns text report with output format text', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'text' } } as any, () => {
      try {
        assert(log[0].indexOf('-----------------------') > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns json report with output format default', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: {} } as any, () => {
      try {
        assert(Array.isArray(log[0]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('writes CodeTour doctor report to .tours folder when in tour output mode. Creates the folder when it does not exist', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return false;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');

    command.action(logger, { options: { output: 'tour' } } as any, () => {
      try {
        assert(writeFileSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours/validation.tour')), 'Tour file not created');
        assert(mkDirSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours')), '.tours folder not created');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('writes CodeTour upgrade report to .tours folder when in tour output mode. Does not create the folder when it already exists', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return true;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');

    command.action(logger, { options: { output: 'tour' } } as any, () => {
      try {
        assert(writeFileSyncStub.calledWith(path.join(process.cwd(), invalidProjectPath, '/.tours/validation.tour')), 'Tour file not created');
        assert(mkDirSyncStub.notCalled, '.tours folder created');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.0.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 7);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.0.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 7);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.0.2 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 7);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.1.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 14);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.1.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 14);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.1.3 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 14);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.2.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 14);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.3.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 15);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.3.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 15);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.3.2 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 15);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.3.4 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 16);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.4.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 13);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.4.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 13);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.5.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.5.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.6.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.7.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.8.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.8.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.8.2 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.9.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.10.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 8);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.11.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.12.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.12.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.13.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.13.1 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-react'));

    command.action(logger, { options: { } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for a valid 1.14.0 project (json)', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), validProjectPath));

    command.action(logger, { options: { output: 'json' } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct message a valid 1.14.0 project (text)', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), validProjectPath));

    command.action(logger, { options: { output: 'text' } } as any, () => {
      try {
        assert.strictEqual(log[0], '✅ CLI for Microsoft 365 has found no issues in your project');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct message for a valid 1.14.0 project (md)', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), validProjectPath));

    command.action(logger, { options: { output: 'md' } } as any, () => {
      try {
        assert(log[0].indexOf('✅ CLI for Microsoft 365 has found no issues in your project') > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows yarn commands for yarn package manager', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'json', packageManager: 'yarn' } } as any, () => {
      try {
        const findings: FindingToReport[] = log.pop();
        assert.strictEqual(findings[0].resolution.indexOf('yarn '), 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows yarn commands for pnpm package manager', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'json', packageManager: 'pnpm' } } as any, () => {
      try {
        const findings: FindingToReport[] = log.pop();
        assert.strictEqual(findings[0].resolution.indexOf('pnpm '), 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for an invalid 1.14.0 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'json' } } as any, () => {
      try {
        const findings: FindingToReport[] = log[0];
        assert.strictEqual(findings.length, 28);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('e2e: shows correct number of findings for an invalid 1.14.0 project (debug)', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), invalidProjectPath));

    command.action(logger, { options: { output: 'json', debug: true } } as any, () => {
      try {
        const findings: FindingToReport[] = log.pop();
        assert.strictEqual(findings.length, 28);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('passes validation when package manager not specified', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported package manager specified', () => {
    const actual = command.validate({ options: { packageManager: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when npm package manager specified', () => {
    const actual = command.validate({ options: { packageManager: 'npm' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when pnpm package manager specified', () => {
    const actual = command.validate({ options: { packageManager: 'pnpm' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when yarn package manager specified', () => {
    const actual = command.validate({ options: { packageManager: 'yarn' } });
    assert.strictEqual(actual, true);
  });
});