import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-set');

describe(commands.APP_SET, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch
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
    assert.strictEqual(command.name, commands.APP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates uri for the specified appId', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates uri for the specified objectId', (done) => {
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates uri for the specified name', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return Promise.resolve();
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips updating uri if no uri specified', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with objectId not found', (done) => {
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));

    command.action(logger, {
      options: {
        debug: false,
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with the appId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with name not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with name My app found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when multiple apps with the specified name found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple Azure AD application registration with name My app found. Please disambiguate (app object IDs): 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appId failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through name failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));
    sinon.stub(request, 'patch').callsFake(_ => Promise.reject('PATCH request executed'));

    command.action(logger, {
      options: {
        debug: false,
        name: 'My app',
        uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if appId and objectId specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', () => {
    const actual = command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId nor name specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', () => {
    const actual = command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)', () => {
    const actual = command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', () => {
    const actual = command.validate({ options: { name: 'My app', uri: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } });
    assert.strictEqual(actual, true);
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
