import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';

const command: Command = require('./contenttypehub-get');

describe(commands.CONTENTTYPEHUB_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    auth.service.connected = true;
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
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPEHUB_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should send correct request body', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      // fake contenttype hub url retrieval
      if (opts.data.indexOf('981cbc68-9edc-4f8d-872f-71146fcbb84f') > -1) {
        return JSON.stringify([{
          "SchemaVersion": "15.0.0.0",
          "LibraryVersion": "16.0.7331.1206",
          "ErrorInfo": null,
          "TraceCorrelationId": "ca54ff9e-8062-1023-19f5-865d949b3748"
        }, 7, {
          "_ObjectType_": "SP.Taxonomy.TermStore",
          "_ObjectIdentity_": "ca54ff9e-8062-1000-18f5-865d949b3748|fec14c62-7c3b-481b-851b-c80d7802b224:st:mY10nDmmVEbNU++TAiFjtQ==",
          "ContentTypePublishingHub": "https:\\u002f\\u002fcontoso.sharepoint.com\\u002fsites\\u002fcontentTypeHub"
        }]);
      }

      throw 'Invalid request';
    });

    const options = {
      verbose: true
    };

    await command.action(logger, { options: options } as any);
    const bodyPayload = `<Request xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\" AddExpandoFieldTypeSuffix=\"true\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\"${config.applicationName}\">\n<Actions>\n  <ObjectPath Id=\"2\" ObjectPathId=\"1\" />\n  <ObjectIdentityQuery Id=\"3\" ObjectPathId=\"1\" />\n  <ObjectPath Id=\"5\" ObjectPathId=\"4\" />\n  <ObjectIdentityQuery Id=\"6\" ObjectPathId=\"4\" />\n  <Query Id=\"7\" ObjectPathId=\"4\">\n    <Query SelectAllProperties=\"false\">\n      <Properties>\n        <Property Name=\"ContentTypePublishingHub\" ScalarProperty=\"true\" />\n      </Properties>\n    </Query>\n  </Query>\n</Actions>\n<ObjectPaths>\n  <StaticMethod Id=\"1\" Name=\"GetTaxonomySession\" TypeId=\"{981cbc68-9edc-4f8d-872f-71146fcbb84f}\" />\n  <Method Id=\"4\" ParentId=\"1\" Name=\"GetDefaultSiteCollectionTermStore\" />\n</ObjectPaths>\n</Request>`;
    assert.strictEqual(requestStub.lastCall.args[0].data, bodyPayload);
    assert(loggerLogSpy.calledWith({ "ContentTypePublishingHub": "https:\\u002f\\u002fcontoso.sharepoint.com\\u002fsites\\u002fcontentTypeHub" }));
  });

  it('should correctly handle reject promise', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.data.indexOf('981cbc68-9edc-4f8d-872f-71146fcbb84f') > -1) {
        throw 'request error';
      }
      throw 'Invalid request';
    });

    const options = {
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('request error'));
  });

  it('should correctly handle ErrorInfo', async () => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "ClientSvc error" } }]);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.data.indexOf('981cbc68-9edc-4f8d-872f-71146fcbb84f') > -1) {
        return error;
      }
      throw 'Invalid request';
    });
    const options = {
      verbose: true
    };
    await assert.rejects(command.action(logger, { options: options } as any),
      new CommandError('ClientSvc error'));
  });

  it('Contains the correct options', () => {
    const options = command.options;
    let containsOutputOption = false;
    let containsVerboseOption = false;
    let containsDebugOption = false;
    let containsQueryOption = false;

    options.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        containsOutputOption = true;
      }
      else if (o.option.indexOf('--verbose') > -1) {
        containsVerboseOption = true;
      }
      else if (o.option.indexOf('--debug') > -1) {
        containsDebugOption = true;
      }
      else if (o.option.indexOf('--query') > -1) {
        containsQueryOption = true;
      }
    });

    assert(options.length === 4, "Wrong amount of options returned");
    assert(containsOutputOption, "Output option not available");
    assert(containsVerboseOption, "Verbose option not available");
    assert(containsDebugOption, "Debug option not available");
    assert(containsQueryOption, "Query option not available");
  });
});
