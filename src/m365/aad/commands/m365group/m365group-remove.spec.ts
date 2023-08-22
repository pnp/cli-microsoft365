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
import { spo } from '../../../../utils/spo.js';
import command from './m365group-remove.js';
import { aadGroup } from '../../../../utils/aadGroup.js';

describe(commands.M365GROUP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  const groupId = "3e6e705d-6fb5-4ca7-84dc-3c8f5154fe2c";

  const defaultGetStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/teams/sales/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}`) {
        return { id: groupId };
      }

      throw 'Invalid request';
    });
  };

  const defaultPostStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/GroupSiteManager/Delete?siteUrl='https://contoso.sharepoint.com/teams/sales'`) {
        return Promise.resolve({
          "data": {
            "odata.null": true
          }
        });
      }

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302" }, 185, { "IsNull": false }]);
      }

      throw 'Invalid request';
    });
  };

  const defaultDeleteStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}`) {
        return;
      }
      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(true);
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.delete,
      spo.getSpoAdminUrl,
      spo.ensureFormDigest,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified group when confirm option not passed', async () => {
    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the group when prompt not confirmed', async () => {
    const getGroupSpy: sinon.SinonStub = defaultGetStub();

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(getGroupSpy.notCalled);
  });

  it('deletes the group site for the sepcified group id when prompt confirmed', async () => {
    defaultGetStub();
    const deletedGroupSpy: sinon.SinonStub = defaultPostStub();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { id: groupId, verbose: true } });
    assert(deletedGroupSpy.calledOnce);
    assert(loggerLogToStderrSpy.calledWith(`Deleting the group site: 'https://contoso.sharepoint.com/teams/sales'...`));
  });

  it('deletes the group without moving it to the Recycle Bin', async () => {
    defaultGetStub();
    defaultPostStub();
    const deleteStub: sinon.SinonStub = defaultDeleteStub();

    await command.action(logger, { options: { id: groupId, verbose: true, skipRecycleBin: true, force: true } });
    assert(deleteStub.called);
    assert(loggerLogToStderrSpy.calledWith("Group has been deleted and is now available in the deleted items list. Removing permanently..."));
  });

  it('verifies if the group is deleted and available in the deleted groups list, retry and delete the group', async () => {
    const getCallStub: sinon.SinonStub = sinon.stub(request, 'get');

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/groups/${groupId}/drive?$select=webUrl` }))
      .resolves({ webUrl: "https://contoso.sharepoint.com/teams/sales/Shared%20Documents" });

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}` }))
      .onFirstCall().rejects({ response: { status: 404 } })
      .onSecondCall().resolves({ id: groupId });

    defaultPostStub();
    const deleteStub: sinon.SinonStub = defaultDeleteStub();

    await command.action(logger, { options: { id: groupId, verbose: true, skipRecycleBin: true, force: true } });
    assert(loggerLogToStderrSpy.calledWith("Group has not been deleted yet. Waiting and retrying..."));
    assert(deleteStub.called);
  });

  it('handles error if unexpected error occurs while finding the group in the deleted groups list', async () => {
    const getCallStub: sinon.SinonStub = sinon.stub(request, 'get');

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/groups/${groupId}/drive?$select=webUrl` }))
      .resolves({ webUrl: "https://contoso.sharepoint.com/teams/sales/Shared%20Documents" });

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}` }))
      .rejects();

    defaultPostStub();
    defaultDeleteStub();

    await assert.rejects(command.action(logger, { options: { id: groupId, verbose: true, skipRecycleBin: true, force: true } }),
      new CommandError('Error'));
  });

  it('handles group not found after all retries', async () => {
    const getCallStub: sinon.SinonStub = sinon.stub(request, 'get');

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/groups/${groupId}/drive?$select=webUrl` }))
      .resolves({ webUrl: "https://contoso.sharepoint.com/teams/sales/Shared%20Documents" });

    getCallStub.withArgs(sinon.match({ url: `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}` }))
      .rejects({ response: { status: 404 } });

    defaultPostStub();
    const deleteStub: sinon.SinonStub = defaultDeleteStub();

    await command.action(logger, { options: { id: groupId, verbose: true, skipRecycleBin: true, force: true } });
    assert(deleteStub.notCalled);
  });

  it('throws error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(aadGroup.isUnifiedGroup);
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId, force: true } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
