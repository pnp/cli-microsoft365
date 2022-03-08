import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-conversation-post-list');

describe(commands.O365GROUP_CONVERSATION_POST_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const jsonOutput = {
    "value": [
      {
        "id": "AAMkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQBGAAAAAAAItFGwjIkpSKk3RMD2kEsABwB8V4aGbsmzQpcmFTaihptDAAAAAAEMAAB8V4aGbsmzQpcmFTaihptDAAAAABUFAAA=",
        "createdDateTime": "2022-02-21T22:13:53Z",
        "lastModifiedDateTime": "2022-02-21T22:13:53Z",
        "changeKey": "CQAAABYAAAB8V4aGbsmzQpcmFTaihptDAAAAAAKN",
        "categories": [],
        "receivedDateTime": "2022-02-21T22:13:53Z",
        "hasAttachments": false,
        "body": {
          "contentType": "html",
          "content": "<html><body><div>\r\\\n<div dir=\"ltr\">\r\\\n<div dir=\"ltr\">\r\\\n<div style=\"color:black;font-size:12pt;font-family:Calibri,Arial,Helvetica,sans-serif;\">\r\\\nThis is one</div>\r\\\n</div>\r\\\n</div>\r\\\n</div>\r\\\n</body></html>"
        },
        "from": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        },
        "sender": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        }
      },
      {
        "id": "AAMkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQBGAAAAAAAItFGwjIkpSKk3RMD2kEsABwB8V4aGbsmzQpcmFTaihptDAAAAAAEMAAB8V4aGbsmzQpcmFTaihptDAAAAABUGAAA=",
        "createdDateTime": "2022-02-21T22:14:14Z",
        "lastModifiedDateTime": "2022-02-21T22:14:14Z",
        "changeKey": "CQAAABYAAAB8V4aGbsmzQpcmFTaihptDAAAAAAKa",
        "categories": [],
        "receivedDateTime": "2022-02-21T22:14:14Z",
        "hasAttachments": false,
        "body": {
          "contentType": "html",
          "content": "<html><body><div>\r\\\n<div dir=\"ltr\">\r\\\n<div style=\"color:black;font-size:12pt;font-family:Calibri,Arial,Helvetica,sans-serif;\">\r\\\nReply to One</div>\r\\\n</div>\r\\\n</div>\r\\\n</body></html>"
        },
        "from": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        },
        "sender": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        }
      }
    ]
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_CONVERSATION_POST_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['receivedDateTime', 'id']);
  });
  it('fails validation if groupId and groupDisplayName specified', () => {
    const actual = command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', groupDisplayName: 'MyGroup' } });
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if neither groupId nor groupDisplayName specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the groupId is not a valid GUID', () => {
    const actual = command.validate({ options: { groupId: 'not-c49b-4fd4-8223-28f0ac3a6402' } });
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the groupId is a valid GUID', () => {
    const actual = command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.strictEqual(actual, true);
  });

  it('Retrieve posts for the specified conversation threadId of o365 group groupId in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/threads/AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E=/posts`) {
        return Promise.resolve(
          jsonOutput
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true,
        groupId: "00000000-0000-0000-0000-000000000000",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutput.value
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('Retrieve posts for the specified conversation threadId of o365 group groupDisplayName in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
            }
          ]
        });
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/threads/AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E=/posts`) {
        return Promise.resolve(
          jsonOutput
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true,
        groupDisplayName: "MyGroup",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutput.value
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('handles error when the group with the specified the name not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'MyGroup'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        groupDisplayName: 'MyGroup',
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `The specified group does not exist`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('handles error when multiple groups with the specified name found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'MyGroup'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '8b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        groupDisplayName: 'MyGroup',
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple groups found with name MyGroup found. Please choose between the following IDs: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 8b1b1e42-794b-4c71-93ac-5ed92488b67g`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly handles error when listing posts', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false,
        groupId: "00000000-0000-0000-0000-000000000000",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
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