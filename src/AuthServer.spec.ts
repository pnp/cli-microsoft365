import * as assert from 'assert';
import Axios from 'axios';
import { AddressInfo } from 'net';
import { Auth } from './Auth';
import authServer from './AuthServer';
import { Logger } from './cli';
import sinon = require('sinon');

describe('AuthServer', () => {
  let log: any[];

  let callbackResolveStub: sinon.SinonStub;
  let callbackRejectStub: sinon.SinonStub;
  let openStub: sinon.SinonStub;
  let serverUrl: string = "";
  let auth: Auth;

  const logger: Logger = {
    log: (msg: any) => log.push(msg),
    logRaw: (msg: any) => log.push(msg),
    logToStderr: (msg: any) => log.push(msg)
  }

  beforeEach(() => {
    log = [];
    auth = new Auth()
    auth.service.appId = '9bc3ab49-b65d-410a-85ad-de819febfddc';
    auth.service.tenant = '9bc3ab49-b65d-410a-85ad-de819febfddd';
    openStub = sinon.stub(authServer as any, 'open').callsFake(_ => Promise.resolve());
    callbackResolveStub = sinon.stub().callsFake(() => { })
    callbackRejectStub = sinon.stub().callsFake(() => { })
    authServer.initializeServer(auth.service, auth.defaultResource, callbackResolveStub, callbackRejectStub, logger, true);
    const address = authServer.server.address() as AddressInfo;
    serverUrl = `http://localhost:${address?.port}`;
  });

  afterEach(() => {
    if (authServer.server.listening) {
      authServer.server.close();
    }
    openStub.restore();
  });

  it('successfully listens', (done) => {
    const server = authServer.server;

    try {
      assert(server !== undefined && server !== null);
      assert(server.listening);
      assert(serverUrl.indexOf("http://localhost:") > -1);
      done();
    }
    catch (err) {
      done(err);
    }
  });

  it('successfully returns an auth code', (done) => {
    const server = authServer.server;

    try {
      assert(server !== undefined && server !== null);
      assert(server.listening);
      assert(serverUrl.indexOf("http://localhost:") > -1);

      Axios.get(`${serverUrl}/?code=1111`).then((response) => {
        assert(response.data.indexOf("You have logged into CLI for Microsoft 365!") > -1);
        assert(callbackResolveStub.called);
        assert(callbackResolveStub.args[0][0].code === "1111");
        assert(callbackResolveStub.args[0][0].redirectUri === serverUrl);
        assert(callbackRejectStub.notCalled);
        assert(authServer.server.listening === false, "server is closed after a successful request");
        done();
      }).catch((reason) => {
        done(reason);
      })
    }
    catch (err) {
      done(err);
    }
  });

  it('successfully returns error message only', (done) => {
    try {
      Axios.get(`${serverUrl}/?error=an error has occurred`).then((response) => {
        assert(response.data.indexOf("Oops! Azure Active Directory replied with an error message.") > -1);
        assert(callbackResolveStub.notCalled);
        assert(callbackRejectStub.called);
        assert(callbackRejectStub.args[0][0].error === "an error has occurred");
        assert(callbackRejectStub.args[0][0].errorDescription === undefined);
        assert(authServer.server.listening === false, "server is closed after a successful request");
        done();
      }).catch((reason) => {
        done(reason);
      })
    }
    catch (err) {
      done(err);
    }
  });

  it('successfully returns error message and error description', (done) => {
    try {
      Axios.get(`${serverUrl}/?error=an error has occurred&error_description=error description`).then((response) => {
        assert(response.data.indexOf("Oops! Azure Active Directory replied with an error message.") > -1);
        assert(callbackResolveStub.notCalled);
        assert(callbackRejectStub.called);
        assert(callbackRejectStub.args[0][0].error === "an error has occurred");
        assert(callbackRejectStub.args[0][0].errorDescription === "error description");
        assert(authServer.server.listening === false, "server is closed after a successful request");
        done();
      }).catch((reason) => {
        done(reason);
      })
    }
    catch (err) {
      done(err);
    }
  });

  it('fails if there is an invalid request', (done) => {
    try {
      Axios.get(`${serverUrl}/?requestingSomthingElse=true`).then((response) => {
        assert(response.data.indexOf("Oops! This is an invalid request.") > -1);
        assert(callbackResolveStub.notCalled);
        assert(callbackRejectStub.called);
        assert(callbackRejectStub.args[0][0].error === "invalid request");
        assert(callbackRejectStub.args[0][0].errorDescription === "An invalid request has been received by the HTTP server");
        assert(authServer.server.listening === false, "server is closed after a successful request");
        done();
      }).catch((reason) => {
        done(reason);
      })
    }
    catch (err) {
      done(err);
    }
  });

  it('fails if open fails', (done) => {
    try {
      if (authServer.server.listening) {
        authServer.server.close();
      }

      openStub.restore();
      openStub = sinon.stub(authServer as any, 'open').callsFake(_ => Promise.reject());
      authServer.initializeServer(auth.service, auth.defaultResource, callbackResolveStub, callbackRejectStub, logger, true);
      setTimeout(() => {
        assert(callbackRejectStub.called);
        assert(callbackRejectStub.args[0][0].error === "Can't open the default browser");
        assert(callbackRejectStub.args[0][0].errorDescription === "Was not able to open a browser instance. Try again later or use a different authentication method.");
        assert(authServer.server.listening === false, "server is closed after a successful request");
        done();
      }, 10);
    }
    catch (err) {
      done(err);
    }
  });
});