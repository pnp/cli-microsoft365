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
const command: Command = require('./group-add');

const validSharePointUrl = 'https://contoso.sharepoint.com/sites/project-x';
const validName = 'Project leaders';

const groupAddedResponse = {
  Id: 1,
  Title: validName,
  AllowMembersEditMembership: false,
  AllowRequestToJoinLeave: false,
  AutoAcceptRequestToJoinLeave: false,
  Description: 'Lorem ipsum',
  OnlyAllowMembersViewMembership: false,
  RequestToJoinLeaveEmailSetting: 'john.doe@contoso.com'
};

describe(commands.GROUP_ADD, () => {
  let log: any[];
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
    assert.strictEqual(command.name, commands.GROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', name: validName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url is valid and name is passed', async () => {
    const actual = await command.validate({ options: { webUrl: validSharePointUrl, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds group to site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validSharePointUrl}/_api/web/sitegroups`) {
        return Promise.resolve(groupAddedResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        webUrl: validSharePointUrl,
        name: validName
      }
    });
    assert(loggerLogSpy.calledWith(groupAddedResponse));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError("An error has occurred."));
  });
}); 
