import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./groupsetting-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.GROUPSETTING_GET, () => {
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
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.GROUPSETTING_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified Group Setting', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Group Setting (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "displayName": "Group Setting",
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "templateId": "bb4f86e1-a598-4101-affc-97c6b136a753",
          "values": [
            {
              "name": "Name1",
              "value": "Value1"
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no group setting found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/1caf7dcd-7e83-4c3a-94f7-932a1299c843`) {
        return Promise.reject({
          error: {
            "error": {
              "code": "Request_ResourceNotFound",
              "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.",
              "innerError": {
                "request-id": "7e192558-7438-46db-a4c9-5dca83d2ec96",
                "date": "2018-02-21T20:38:50"
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c843' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.strictEqual(actual, true);
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});