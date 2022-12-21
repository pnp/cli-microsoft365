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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./retentionlabel-add');

describe(commands.RETENTIONLABEL_ADD, () => {
  const invalid = 'invalid';
  const displayName = 'some label';
  const behaviorDuringRetentionPeriod = 'retain';
  const actionAfterRetentionPeriod = 'delete';
  const retentionDuration = 365;
  const retentionTrigger = 'dateLabeled';
  const defaultRecordBehavior = 'startLocked';
  const descriptionForUsers = 'Description for users';
  const descriptionForAdmins = 'Description for admins';
  const labelToBeApplied = 'another label';

  const requestResponse = {
    displayName: "some label",
    descriptionForAdmins: "Description for admins",
    descriptionForUsers: "Description for users",
    isInUse: false,
    retentionTrigger: "dateLabeled",
    behaviorDuringRetentionPeriod: "retain",
    actionAfterRetentionPeriod: "delete",
    createdDateTime: "2022-12-21T09:28:37Z",
    lastModifiedDateTime: "2022-12-21T09:28:37Z",
    labelToBeApplied: "another label",
    defaultRecordBehavior: "startLocked",
    id: "f7e05955-210b-4a8e-a5de-3c64cfa6d9be",
    retentionDuration: {
      days: 365
    },
    createdBy: {
      user: {
        id: null,
        displayName: "John Doe"
      }
    },
    lastModifiedBy: {
      user: {
        id: null,
        displayName: "John Doe"
      }
    },
    dispositionReviewStages: []
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.RETENTIONLABEL_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if retentionDuration is not a number', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: invalid
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration,
        retentionTrigger: retentionTrigger,
        defaultRecordBehavior: defaultRecordBehavior,
        descriptionForUsers: descriptionForUsers,
        descriptionForAdmins: descriptionForAdmins,
        labelToBeApplied: labelToBeApplied
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid behaviorDuringRetentionPeriod', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: invalid,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration
      }
    }, commandInfo);
    assert.strictEqual(actual, `${invalid} is not a valid behavior of a document with the label. Allowed values are doNotRetain|retain|retainAsRecord|retainAsRegulatoryRecord`);
  });

  it('rejects invalid actionAfterRetentionPeriod', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: invalid,
        retentionDuration: retentionDuration
      }
    }, commandInfo);
    assert.strictEqual(actual, `${invalid} is not a valid action to take on a document with the label. Allowed values are none|delete|startDispositionReview`);
  });

  it('rejects invalid retentionTrigger', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration,
        retentionTrigger: invalid
      }
    }, commandInfo);
    assert.strictEqual(actual, `${invalid} is not a valid action retention duration calculation. Allowed values are dateLabeled|dateCreated|dateModified|dateOfEvent`);
  });

  it('rejects invalid defaultRecordBehavior', async () => {
    const actual = await command.validate({
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration,
        defaultRecordBehavior: invalid
      }
    }, commandInfo);
    assert.strictEqual(actual, `${invalid} is not a valid state of a record label. Allowed values are startLocked|startUnlocked`);
  });

  it('Adds retention label by id when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        return requestResponse;
      }

      return 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration,
        retentionTrigger: retentionTrigger,
        defaultRecordBehavior: defaultRecordBehavior,
        descriptionForUsers: descriptionForUsers,
        descriptionForAdmins: descriptionForAdmins,
        labelToBeApplied: labelToBeApplied
      }
    });

    assert(loggerLogSpy.calledWith(requestResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        displayName: displayName,
        behaviorDuringRetentionPeriod: behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: actionAfterRetentionPeriod,
        retentionDuration: retentionDuration
      }
    }), new CommandError("An error has occurred"));
  });
});