import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./funsettings-set');

describe(commands.FUNSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FUNSETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets allowGiphy settings to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: 'false' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowGiphy settings to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: 'true' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets giphyContentRating to moderate', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'moderate'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'moderate' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets giphyContentRating to strict', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'strict'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'strict' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowStickersAndMemes to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: 'true' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowStickersAndMemes to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: 'false' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('sets allowCustomMemes to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'true' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCustomMemes to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'false' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCustomMemes to false (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'false' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'patch').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, {
      options: {
        debug: true,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        allowGiphy: true,
        giphyContentRating: "moderate",
        allowStickersAndMemes: false,
        allowCustomMemes: true
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: { teamId: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when teamId is a valid GUID', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when allowGiphy is a valid boolean', async () => {
    const actualTrue = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'true' }
    }, commandInfo);

    const actualFalse = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'false' }
    }, commandInfo);

    const actual = actualTrue && actualFalse;
    assert.strictEqual(actual, true);
  });

  it('fails validation when allowGiphy is not a valid boolean', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'trueish' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when giphyContentRating is moderate or strict', async () => {
    const actualModerate = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'moderate' }
    }, commandInfo);

    const actualStrict = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'strict' }
    }, commandInfo);

    const actual = actualModerate && actualStrict;
    assert.strictEqual(actual, true);
  });

  it('fails validation when giphyContentRating is not moderate or strict', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'somethingelse' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when allowStickersAndMemes is a valid boolean', async () => {
    const actualTrue = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'true' }
    }, commandInfo);

    const actualFalse = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'false' }
    }, commandInfo);

    const actual = actualTrue && actualFalse;
    assert.strictEqual(actual, true);
  });

  it('fails validation when allowStickersAndMemes is not a valid boolean', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'somethingelse' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when allowCustomMemes is a valid boolean', async () => {
    const actualTrue = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'true' }
    }, commandInfo);

    const actualFalse = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'false' }
    }, commandInfo);

    const actual = actualTrue && actualFalse;
    assert.strictEqual(actual, true);
  });

  it('fails validation when allowCustomMemes is not a valid boolean', async () => {
    const actual = await command.validate({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'somethingelse' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});