import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-reply-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_MESSAGE_REPLY_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      action: command.action(),
      commandWrapper: {
        command: command.name
      },
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
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
    assert.equal(command.name.startsWith(commands.TEAMS_MESSAGE_REPLY_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_MESSAGE_REPLY_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('fails validation if teamId, channelId and messageId are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if channelId and messageId are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if messageId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "5f5d7b71-1161-44",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    });
    assert.notEqual(actual, true);
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

  it('validates for a correct input', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    });
    assert.equal(actual, true);
  });

  it('fails validates for a incorrect channelId missing leading 19:.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        messageId: "1501527481624"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validates for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
        messageId: "1501527481624"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('retrieves the replies to the specified message (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315/channels/19:d0bba23c2fc8413991125a43a54cc30e@thread.skype/messages/1501527481624/replies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('02bd9fd6-8f93-4758-87c3-1fb73740a315')/channels('19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype')/messages('1501527481624')/replies",
          "@odata.count": 2,
          value: [
            {
              "id": "1501527483334",
              "replyToId": "1501527481624",
              "etag": "1501527483334",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:03.334Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527483334?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527483334&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "2ed03dfd-01d8-4005-a9ef-fa8ee546dc6c",
                  "displayName": "Lidia Holloway",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            },
            {
              "id": "1501527482612",
              "replyToId": "1501527481624",
              "etag": "1501527482612",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:02.612Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527482612?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527482612&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "8b209ac8-08ff-4ef1-896d-3b9fde0bbf04",
                  "displayName": "Joni Sherman",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1501527483334",
            "body": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
          },
          {
            "id": "1501527482612",
            "body": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the replies to the specified message', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315/channels/19:d0bba23c2fc8413991125a43a54cc30e@thread.skype/messages/1501527481624/replies`) {
        return Promise.resolve({
          value: [
            {
              "id": "1501527483334",
              "replyToId": "1501527481624",
              "etag": "1501527483334",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:03.334Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527483334?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527483334&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "2ed03dfd-01d8-4005-a9ef-fa8ee546dc6c",
                  "displayName": "Lidia Holloway",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            },
            {
              "id": "1501527482612",
              "replyToId": "1501527481624",
              "etag": "1501527482612",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:02.612Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527482612?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527482612&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "8b209ac8-08ff-4ef1-896d-3b9fde0bbf04",
                  "displayName": "Joni Sherman",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1501527483334",
            "body": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
          },
          {
            "id": "1501527482612",
            "body": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all data in json output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315/channels/19:d0bba23c2fc8413991125a43a54cc30e@thread.skype/messages/1501527481624/replies`) {
        return Promise.resolve({
          value: [
            {
              "id": "1501527483334",
              "replyToId": "1501527481624",
              "etag": "1501527483334",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:03.334Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527483334?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527483334&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "2ed03dfd-01d8-4005-a9ef-fa8ee546dc6c",
                  "displayName": "Lidia Holloway",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            },
            {
              "id": "1501527482612",
              "replyToId": "1501527481624",
              "etag": "1501527482612",
              "messageType": "message",
              "createdDateTime": "2017-07-31T18:58:02.612Z",
              "lastModifiedDateTime": null,
              "deletedDateTime": null,
              "subject": "",
              "summary": null,
              "importance": "normal",
              "locale": "en-us",
              "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527482612?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527482612&parentMessageId=1501527481624",
              "policyViolation": null,
              "from": {
                "application": null,
                "device": null,
                "conversation": null,
                "user": {
                  "id": "8b209ac8-08ff-4ef1-896d-3b9fde0bbf04",
                  "displayName": "Joni Sherman",
                  "userIdentityType": "aadUser"
                }
              },
              "body": {
                "contentType": "html",
                "content": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
              },
              "attachments": [],
              "mentions": [],
              "reactions": []
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "1501527483334",
            "replyToId": "1501527481624",
            "etag": "1501527483334",
            "messageType": "message",
            "createdDateTime": "2017-07-31T18:58:03.334Z",
            "lastModifiedDateTime": null,
            "deletedDateTime": null,
            "subject": "",
            "summary": null,
            "importance": "normal",
            "locale": "en-us",
            "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527483334?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527483334&parentMessageId=1501527481624",
            "policyViolation": null,
            "from": {
              "application": null,
              "device": null,
              "conversation": null,
              "user": {
                "id": "2ed03dfd-01d8-4005-a9ef-fa8ee546dc6c",
                "displayName": "Lidia Holloway",
                "userIdentityType": "aadUser"
              }
            },
            "body": {
              "contentType": "html",
              "content": "<div>Hey team, I'm Lidia! I've been here about six months so far and I really like it! We've got a great team and although there's always so much to do, I enjoy how well we work together.</div>"
            },
            "attachments": [],
            "mentions": [],
            "reactions": []
          },
          {
            "id": "1501527482612",
            "replyToId": "1501527481624",
            "etag": "1501527482612",
            "messageType": "message",
            "createdDateTime": "2017-07-31T18:58:02.612Z",
            "lastModifiedDateTime": null,
            "deletedDateTime": null,
            "subject": "",
            "summary": null,
            "importance": "normal",
            "locale": "en-us",
            "webUrl": "https://teams.microsoft.com/l/message/19%3Ad0bba23c2fc8413991125a43a54cc30e%40thread.skype/1501527482612?groupId=02bd9fd6-8f93-4758-87c3-1fb73740a315&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35&createdTime=1501527482612&parentMessageId=1501527481624",
            "policyViolation": null,
            "from": {
              "application": null,
              "device": null,
              "conversation": null,
              "user": {
                "id": "8b209ac8-08ff-4ef1-896d-3b9fde0bbf04",
                "displayName": "Joni Sherman",
                "userIdentityType": "aadUser"
              }
            },
            "body": {
              "contentType": "html",
              "content": "<div>Hi everyone, I'm Joni and I've been with our group for about 6 years. Feel free to ping me with any questions you may have!</div>"
            },
            "attachments": [],
            "mentions": [],
            "reactions": []
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving replies', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        channelId: "19:d0bba23c2fc8413991125a43a54cc30e@thread.skype",
        messageId: "1501527481624"
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});