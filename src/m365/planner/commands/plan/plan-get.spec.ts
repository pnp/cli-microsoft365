import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./plan-get');

describe(commands.PLAN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const validId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTitle = 'Plan name';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const invalidOwnerGroupId = 'Invalid GUID';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const planResponse = {
    "id": validId,
    "title": validTitle
  };

  const planDetailsResponse = {
    "sharedWith": { },
    "categoryDescriptions": { }
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PLAN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner', '@odata.etag']);
  });
  
  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('fails validation if neither id nor title are provided.', (done) => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both id and title are specified', (done) => {
    const actual = command.validate({
      options: {
        id: validId,
        title: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both deprecated planId and planTitle are specified', (done) => {
    const actual = command.validate({
      options: {
        planId: validId,
        planTitle: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both id and deprecated planTitle are specified', (done) => {
    const actual = command.validate({
      options: {
        id: validId,
        planTitle: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both title and deprecated planId are specified', (done) => {
    const actual = command.validate({
      options: {
        planId: validId,
        title: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided.', (done) => {
    const actual = command.validate({
      options: {
        title: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        title: validTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided with deprecated planTitle', (done) => {
    const actual = command.validate({
      options: {
        planTitle: validTitle
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified with deprecated planTitle', (done) => {
    const actual = command.validate({
      options: {
        planTitle: validTitle,
        ownerGroupId: validOwnerGroupId,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when id specified', (done) => {
    const actual = command.validate({
      options: {
        id: validId
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when title and valid ownerGroupId specified', (done) => {
    const actual = command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when title and valid ownerGroupName specified', (done) => {
    const actual = command.validate({
      options: {
        title: validTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly get planner plan with given id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return Promise.resolve(planResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      id: validId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan with deprecated planId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return Promise.resolve(planResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      planId: validId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan with given title and ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan with given ownerGroupId and deprecated planTitle', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      planTitle: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when using app only access token', (done) => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    command.action(logger, {
      options: {
        id: validId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This command does not support application permissions.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan with given title and ownerGroupName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve(singleGroupResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupName: validOwnerGroupName
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no plan found with given ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
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
});