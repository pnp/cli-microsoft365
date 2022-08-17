import { sinonUtil } from './../../../../utils/sinonUtil';
import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import commands from '../../commands';
import request from '../../../../request';
const command: Command = require('./changelog-list');

describe(commands.CHANGELOG_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const validVersions = 'beta,v1.0';
  const validChangeType = 'Addition';
  const validServices = 'Groups,Security';
  const validStartDate = '2018-12-01';
  const validEndDate = '2019-03-01';

  const validRSSResponse = `
    <rss version="2.0">
      <channel
        xmlns:atom="http://www.w3.org/2005/Atom">
        <title>Microsoft Graph Changelog</title>
        <link>https://graph.office.net/en-us/graph/changelog/rss</link>
        <description>Microsoft Graph Changelog Rss Feed</description>
        <lastBuildDate>Tue, 12 Jul 2022 09:58:42 Z</lastBuildDate>
        <atom:link href="https://graph.office.net/en-us/graph/changelog/rss/?search=&amp;filterBy=Financials" rel="self" type="application/rss+xml" />
        <item>
          <guid isPermaLink="false">7f1afeea-1c73-4e84-af08-8c9cd0fe27d5v1.0</guid>
          <category>prd</category>
          <category>v1.0</category>
          <title>Groups</title>
          <description>Added something.</description>
          <pubDate>2019-01-01T00:00:00.000Z</pubDate>
        </item>
        <item>
          <guid isPermaLink="false">7f1afeea-1c73-4e84-af08-8c9cd0fe27d5beta</guid>
          <category>prd</category>
          <category>beta</category>
          <title>Security</title>
          <description>Added something.</description>
          <pubDate>2019-02-01T00:00:00.000Z</pubDate>
        </item>
      </channel>
    </rss>
  `;

  const validChangelog = [
    {
      guid: '7f1afeea-1c73-4e84-af08-8c9cd0fe27d5beta',
      category: 'beta',
      title: 'Security',
      description: 'Added something.',
      pubDate: new Date('2019-02-01T00:00:00.000Z')
    },
    {
      guid: '7f1afeea-1c73-4e84-af08-8c9cd0fe27d5v1.0',
      category: 'v1.0',
      title: 'Groups',
      description: 'Added something.',
      pubDate: new Date('2019-01-01T00:00:00.000Z')
    }
  ];

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

    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([request.get]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHANGELOG_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['guid', 'category', 'title', 'description', 'pubDate']);
  });

  it('fails validation if versions contains an invalid value.', (done) => {
    const actual = command.validate({
      options: {
        versions: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if changeType is an invalid value.', (done) => {
    const actual = command.validate({
      options: {
        changeType: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if services contains an invalid value.', (done) => {
    const actual = command.validate({
      options: {
        services: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if startDate is invalid ISO date.', (done) => {
    const actual = command.validate({
      options: {
        startDate: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if endDate is invalid ISO date.', (done) => {
    const actual = command.validate({
      options: {
        endDate: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid versions specified', (done) => {
    const actual = command.validate({
      options: {
        versions: validVersions
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid changeType specified', (done) => {
    const actual = command.validate({
      options: {
        changeType: validChangeType
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid services specified', (done) => {
    const actual = command.validate({
      options: {
        services: validServices
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid startDate specified', (done) => {
    const actual = command.validate({
      options: {
        startDate: validStartDate
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid endDate specified', (done) => {
    const actual = command.validate({
      options: {
        endDate: validEndDate
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('retrieves changelog list', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://developer.microsoft.com/en-us/graph/changelog/rss') {
        return Promise.resolve(validRSSResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: { }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(validChangelog));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves changelog list based on changeType', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://developer.microsoft.com/en-us/graph/changelog/rss/?filterBy=Addition') {
        return Promise.resolve(validRSSResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: { 
        changeType: validChangeType
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(validChangelog));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves changelog list based on versions, services, startDate and endDate', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://developer.microsoft.com/en-us/graph/changelog/rss') {
        return Promise.resolve(validRSSResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: { 
        versions: validVersions,
        services: validServices,
        startDate: validStartDate,
        endDate: validEndDate
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(validChangelog));
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

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});