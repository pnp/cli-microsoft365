import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './people-profilecardproperty-remove.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PEOPLE_PROFILECARDPROPERTY_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not a valid value.', async () => {
    const actual = await command.validate({ options: { name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the name is set to userPrincipalName.', async () => {
    const actual = await command.validate({ options: { name: 'userPrincipalName' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly removes profile card property for userPrincipalName', async () => {
    const removeStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'userPrincipalName' } });
    assert(removeStub.called);
  });

  it('correctly removes profile card property for userPrincipalName (debug)', async () => {
    const removeStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'userPrincipalName', debug: true } });
    assert(removeStub.called);
  });

  it('correctly removes profile card property for fax', async () => {
    const removeStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/Fax`) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'fax' } });
    assert(removeStub.called);
  });

  it('correctly removes profile card property for state with force', async () => {
    const removeStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/StateOrProvince`) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'StateOrProvince', force: true } });
    assert(removeStub.called);
  });

  it('uses correct casing for name when incorrect casing is used', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/StateOrProvince`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: 'STATEORPROVINCE', force: true } });
    assert(deleteStub.called);
  });

  it('fails when the removal runs into a property that is not found', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
        throw {
          "error": {
            "code": "404",
            "message": "Not Found",
            "innerError": {
              "peopleAdminErrorCode": "PeopleAdminItemNotFound",
              "peopleAdminRequestId": "2497e6f6-cd91-8bd8-5c53-361d355a5c41",
              "peopleAdminClientRequestId": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276",
              "date": "2023-11-02T19:31:25",
              "request-id": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276",
              "client-request-id": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276"
            }
          }
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'userPrincipalName'
      }
    }), new CommandError(`Not Found`));
  });
});