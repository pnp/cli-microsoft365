import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as fs from 'fs';
const command: Command = require('./message-export');

describe(commands.MESSAGE_EXPORT, () => {
  const userId = '11f43044-095e-456a-b339-7e1901b0c3ae';
  const teamId = '75619fc7-5dce-412b-82ee-f76988d3efaa';
  const fromDateTime = '2023-04-01T00:00:00Z';
  const toDateTime = '2023-04-30T23:59:59Z';
  const folderPath = 'C:\\Temp';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_EXPORT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid userPrincipalName', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userName: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, teamId: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if fromDateTime is not a valid ISO DateTime', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, fromDateTime: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if toDateTime is not a valid ISO DateTime', async () => {
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, toDateTime: 'invalid', withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if folderPath does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, withAttachments: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

  it('passes validation if folderPath exists and userId is a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { folderPath: folderPath, userId: userId, withAttachments: false } }, commandInfo);
    assert.strictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

  it('passes validation if folderPath exists, teamId is a valid GUID and both dates are valid ISO dates', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { folderPath: folderPath, teamId: teamId, fromDateTime: fromDateTime, toDateTime: toDateTime, withAttachments: false } }, commandInfo);
    assert.strictEqual(actual, true);
    sinonUtil.restore(fs.existsSync);
  });

});
