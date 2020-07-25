import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-clone');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import request from '../../../../request';

describe(commands.TEAMS_TEAM_CLONE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TEAM_CLONE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid GUID.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation on invalid visibility', () => {
    const actual = (command.validate() as CommandValidate)({ options: { visibility: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'private\' visibility', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'private'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'public\' visibility', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'public'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct with mandatory parameters', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct with mandatory and optional parameters', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        description: "Self help community for library",
        visibility: "public",
        classification: "public"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "abc"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if partsToClone is set to invalid value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "abc"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "private"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "private"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('creates a clone of a Microsoft Teams team with mandatory parameters', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a clone of a Microsoft Teams team with mandatory parameters (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a clone of a Microsoft Teams team with optional parameters (debug)', (done) => {
    const sinonStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: 'Library Assist',
        partsToClone: 'apps,tabs,settings,channels,members',
        description: 'abc',
        visibility: 'public',
        classification: 'label'
      }
    }, () => {
      try {
        assert.strictEqual(sinonStub.lastCall.args[0].url, 'https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone');
        assert.strictEqual(sinonStub.lastCall.args[0].body.displayName, 'Library Assist');
        assert.strictEqual(sinonStub.lastCall.args[0].body.partsToClone, 'apps,tabs,settings,channels,members');
        assert.strictEqual(sinonStub.lastCall.args[0].body.description, 'abc');
        assert.strictEqual(sinonStub.lastCall.args[0].body.visibility, 'public');
        assert.strictEqual(sinonStub.lastCall.args[0].body.classification, 'label');
        assert.notStrictEqual(sinonStub.lastCall.args[0].body.mailNickname.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({
      options: {
        debug: true,
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: 'Library Assist',
        partsToClone: 'apps,tabs,settings,channels,members',
        description: 'abc',
        visibility: 'public',
        classification: 'label'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});