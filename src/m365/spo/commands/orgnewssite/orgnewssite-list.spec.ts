import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './orgnewssite-list.js';

describe(commands.ORGNEWSSITE_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ORGNEWSSITE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('returns a result', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 22, ['http://contoso.sharepoint.com/sites/site1']
        ]));
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert.strictEqual(svcListRequest.callCount, 1);
    assert(loggerLogSpy.calledWith(['http://contoso.sharepoint.com/sites/site1']));
  });

  it('returns array of results', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          { "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "8992299e-a003-4000-7686-fda36e26a53c" }, 22, ['http://contoso.sharepoint.com/sites/site1', 'http://contoso.sharepoint.com/sites/site2']
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: false } });
    assert.strictEqual(svcListRequest.callCount, 1);
    assert(loggerLogSpy.calledWith(['http://contoso.sharepoint.com/sites/site1', 'http://contoso.sharepoint.com/sites/site2']));
  });

  it('handles error getting request', async () => {
    const svcListRequest = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest']) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
              }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true
      }
    } as any), new CommandError('An error has occurred'));
    assert(svcListRequest.called);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
