import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from "../../commands";
const command: Command = require('./serviceannouncement-message-get');

describe(commands.SERVICEANNOUNCEMENT_MESSAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const testId = 'MC001337';
  const testIncorrectId = '123456';

  const resResourceNotExist = {
    "error": {
      "code": "UnknownError",
      "message": "{\"code\":\"forbidden\",\"message\":\"{\\u0022error\\u0022:\\u0022Resource doesn\\\\u0027t exist for the tenant. ActivityId: b2307a39-e878-458b-bc90-03bc578531d6. Learn more: https://docs.microsoft.com/en-us/graph/api/resources/service-communications-api-overview?view=graph-rest-beta\\\\u0026preserve-view=true.\\u0022}\"}",
      "innerError": {
        "date": "2022-01-22T15:01:15",
        "request-id": "b2307a39-e878-458b-bc90-03bc578531d6",
        "client-request-id": "b2307a39-e878-458b-bc90-03bc578531d6"
      }
    }
  };

  const resMessage = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#admin/serviceAnnouncement/messages/$entity",
    "startDateTime": "2021-02-01T19:23:04Z",
    "endDateTime": "2022-01-31T08:00:00Z",
    "lastModifiedDateTime": "2021-02-01T19:24:37.837Z",
    "title": "Service reminder: Skype for Business Online retires in 6 months",
    "id": "MC001337",
    "category": "planForChange",
    "severity": "normal",
    "tags": [
      "User impact",
      "Admin impact"
    ],
    "isMajorChange": false,
    "actionRequiredByDateTime": "2021-07-31T07:00:00Z",
    "services": [
      "Skype for Business"
    ],
    "expiryDateTime": null,
    "hasAttachments": false,
    "viewPoint": null,
    "details": [
      {
        "name": "BlogLink",
        "value": "https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531"
      },
      {
        "name": "ExternalLink",
        "value": "https://docs.microsoft.com/microsoftteams/skype-for-business-online-retirement"
      }
    ],
    "body": {
      "contentType": "html",
      "content": "<p>Originally announced in MC219641 (July '20), as Microsoft Teams has become the core communications client for Microsoft 365, this is a reminder the Skype for Business Online service will <a href=\"https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531\" target=\"_blank\">retire July 31, 2021</a>. At that point, access to the service will end.</p><p>Please click Additional Information to learn more.</p>"
    }
  };

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
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEANNOUNCEMENT_MESSAGE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if incorrect message ID is provided', (done) => {
    const actual = command.validate({
      options: {
        id: testIncorrectId
      }
    });
    assert.strictEqual(actual, `${testIncorrectId} is not a valid message ID`);
    done();
  });

  it('passes validation if correct message ID is provided', (done) => {
    const actual = command.validate({
      options: {
        id: testId
      }
    });
    assert(actual);
    done();
  });

  it('correctly retrieves service update message', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return Promise.resolve(resMessage);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        id: testId
      }
    }, () => {
      try {
        assert.strictEqual(loggerLogSpy.calledWith(resMessage), true);
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, testId);
        assert.strictEqual(loggerLogSpy.callCount, 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly retrieves service update message (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return Promise.resolve(resMessage);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        id: testId
      }
    }, () => {
      try {
        assert.strictEqual(loggerLogSpy.calledWith(resMessage), true);
        assert.strictEqual(loggerLogSpy.lastCall.args[0].id, testId);
        assert.strictEqual(loggerLogSpy.callCount, 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when the message does not exist for the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testIncorrectId}`) {
        return Promise.reject(resResourceNotExist);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        id: testIncorrectId
      }
    }, (err?: any) => {
      try {
        assert((JSON.parse(JSON.parse(err.message).message).error as string).indexOf(`Resource doesn't exist for the tenant.`) > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    }
    );
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return Promise.resolve(resMessage);
      }

      return Promise.reject('Invalid request');
    });


    command.action(logger, {
      options:
      {
        id: testId,
        output: 'json'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(resMessage));
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