import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';

const command: Command = require('./contenttypehub-get');

describe(commands.CONTENTTYPEHUB_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let stubAllPostRequests: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';

    stubAllPostRequests = (
      contentTypeHubRetrievalResp: any = null
    ): sinon.SinonStub => {
      return sinon.stub(request, 'post').callsFake((opts) => {
        // fake contenttype hub url retrieval
        if (opts.data.indexOf('981cbc68-9edc-4f8d-872f-71146fcbb84f') > -1) {
          if (contentTypeHubRetrievalResp) {
            return contentTypeHubRetrievalResp;
          }
          else {
            return Promise.resolve(JSON.stringify([{
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7331.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "ca54ff9e-8062-1023-19f5-865d949b3748"
            }, 7, {
              "_ObjectType_": "SP.Taxonomy.TermStore",
              "_ObjectIdentity_": "ca54ff9e-8062-1000-18f5-865d949b3748|fec14c62-7c3b-481b-851b-c80d7802b224:st:mY10nDmmVEbNU++TAiFjtQ==",
              "ContentTypePublishingHub": "https:\\u002f\\u002fcontoso.sharepoint.com\\u002fsites\\u002fcontentTypeHub"
            }]));
          }
        }

        return Promise.reject('Invalid request');
      });
    };
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
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPEHUB_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should send correct request body', (done) => {
    const requestStub: sinon.SinonStub = stubAllPostRequests();
    const options = {
      verbose: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        const bodyPayload = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}">
  <Actions>
    <ObjectPath Id="2" ObjectPathId="1" />
    <ObjectIdentityQuery Id="3" ObjectPathId="1" />
    <ObjectPath Id="5" ObjectPathId="4" />
    <ObjectIdentityQuery Id="6" ObjectPathId="4" />
    <Query Id="7" ObjectPathId="4">
      <Query SelectAllProperties="false">
        <Properties>
          <Property Name="ContentTypePublishingHub" ScalarProperty="true" />
        </Properties>
      </Query>
    </Query>
  </Actions>
  <ObjectPaths>
    <StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
    <Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" />
  </ObjectPaths>
</Request>`;
        assert.strictEqual(requestStub.lastCall.args[0].data, bodyPayload);
        assert(loggerLogSpy.calledWith({ "ContentTypePublishingHub": "https:\\u002f\\u002fcontoso.sharepoint.com\\u002fsites\\u002fcontentTypeHub" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle reject promise', (done) => {
    stubAllPostRequests(new Promise<any>((resolve, reject) => { return reject('request error'); }));
    const options = {
      verbose: true
    };
    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('request error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle ErrorInfo', (done) => {
    const error = JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "ClientSvc error" } }]);
    stubAllPostRequests(new Promise<any>((resolve) => { return resolve(error); }));
    const options = {
      verbose: true
    };
    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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