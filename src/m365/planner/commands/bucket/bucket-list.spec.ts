import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './bucket-list.js';

describe(commands.BUCKET_LIST, () => {
  const bucketListResponseValue = [
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket A",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585768731950308408",
      "id": "FtzysDykv0-9s9toWiZhdskAD67z"
    },
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket 2",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585784565[8",
      "id": "ZpcnnvS9ZES2pb91RPxQx8kAMLo5"
    }
  ];

  const bucketListResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerBucket)",
    "@odata.count": 2,
    "value": bucketListResponseValue
  };

  const groupByDisplayNameResponse: any = {
    "value": [
      {
        "id": "0d0402ee-970f-4951-90b5-2f24519d2e40"
      }
    ]
  };

  const plansInOwnerGroup: any = {
    "value": [
      {
        "title": "My Planner Plan",
        "id": "iVPMIgdku0uFlou-KLNg6MkAE1O2"
      },
      {
        "title": "Sample Plan",
        "id": "uO1bj3fdekKuMitpeJqaj8kADBxO"
      }
    ]
  };

  const planResponse = {
    value: [{
      id: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
    }]
  };


  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/RuY-PSpdw02drevnYDTCJpgAEfoI/plans?$select=id`) {
        return planResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'&$select=id`) {
        return groupByDisplayNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans?$select=id,title`) {
        return plansInOwnerGroup;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) {
        return bucketListResponse;
      }
      throw 'Invalid request';
    });
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BUCKET_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'planId', 'orderHint']);
  });

  it('passes validation when valid planId is specified', async () => {
    const actual = await command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid planTitle and ownerGroupId are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid planTitle and ownerGroupName are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly lists planner buckets with planId', async () => {
    const options: any = {
      planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(bucketListResponseValue));
  });

  it('correctly lists planner buckets with planTitle and ownerGroupName', async () => {
    const options: any = {
      planTitle: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(bucketListResponseValue));
  });

  it('correctly lists planner buckets with planTitle and ownerGroupId', async () => {
    const options: any = {
      planTitle: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(bucketListResponseValue));
  });

  it('correctly lists planner buckets by rosterId', async () => {
    const options: any = {
      rosterId: 'RuY-PSpdw02drevnYDTCJpgAEfoI'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(bucketListResponseValue));
  });

  it('fails validation when ownerGroupName not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }), new CommandError(`The specified group 'foo' does not exist.`));
  });

  it('correctly handles API OData error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error("An error has occurred."));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError("An error has occurred."));
  });
});
