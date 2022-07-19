import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { sinonUtil } from '../../../../utils';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', name: validName } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid boolean is passed as option', () => {
    const actual = command.validate({ options: { webUrl: validSharePointUrl, name: validName, allowRequestToJoinLeave: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url is valid and name is passed', () => {
    const actual = command.validate({ options: { webUrl: validSharePointUrl, name: validName } });
    assert.strictEqual(actual, true);
  });

  it('correctly adds group to site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validSharePointUrl}/_api/web/sitegroups`) {
        return Promise.resolve(groupAddedResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        webUrl: validSharePointUrl,
        name: validName
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(groupAddedResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
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