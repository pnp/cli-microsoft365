import { ErrorResponse } from 'adal-node';
import * as http from 'http';
import { IncomingMessage, ServerResponse } from 'http';
import { AddressInfo } from 'net';
import * as open from 'open';
import { ParsedUrlQuery } from 'querystring';
import * as url from "url";
import { InteractiveAuthorizationCodeResponse, Service } from './Auth';
import { Logger } from './cli';

export class AuthServer {
  // assigned through this.initializeServer() hence !
  private httpServer!: http.Server;
  // assigned through this.initializeServer() hence !
  private service!: Service;
  // assigned through this.initializeServer() hence !
  private resolve!: (error: InteractiveAuthorizationCodeResponse) => void;
  // assigned through this.initializeServer() hence !
  private reject!: (error: ErrorResponse) => void;
  // assigned through this.initializeServer() hence !
  private logger!: Logger;

  private open = open;
  private debug: boolean = false;
  private resource: string = "";
  private generatedServerUrl: string = "";

  public get server(): http.Server {
    return this.httpServer;
  }

  public initializeServer = (service: Service, resource: string, resolve: (error: InteractiveAuthorizationCodeResponse) => void, reject: (error: ErrorResponse) => void, logger: Logger, debug: boolean = false) => {
    this.service = service;
    this.resolve = resolve;
    this.reject = reject;
    this.logger = logger;
    this.debug = debug;
    this.resource = resource;

    this.httpServer = http.createServer(this.httpRequest).listen(0, this.httpListener);
  }

  private httpListener = () => {
    const requestState = Math.random().toString(16).substr(2, 20);
    const address = this.httpServer.address() as AddressInfo;
    this.generatedServerUrl = `http://localhost:${address.port}`;
    const url = `https://login.microsoftonline.com/${this.service.tenant}/oauth2/authorize?response_type=code&client_id=${this.service.appId}&redirect_uri=${this.generatedServerUrl}&state=${requestState}&resource=${this.resource}&prompt=select_account`;
    if (this.debug) {
      this.logger.logToStderr('Redirect URL:');
      this.logger.logToStderr(url);
      this.logger.logToStderr('');
    }
    this.openUrl(url);
  }

  private openUrl(url: string) {
    this.open(url).then(() => {
      this.logger.logToStderr("To sign in, use the web browser that just has been opened. Please sign-in there.")
    }).catch(() => {
      const errorResponse: ErrorResponse = {
        error: "Can't open the default browser",
        errorDescription: "Was not able to open a browser instance. Try again later or use a different authentication method."
      }

      this.reject(errorResponse);
      this.httpServer.close();
    });
  }

  private httpRequest = (request: IncomingMessage, response: ServerResponse) => {
    if (this.debug) {
      this.logger.logToStderr('Response:');
      this.logger.logToStderr(request.url);
      this.logger.logToStderr('');
    }

    // url.parse is deprecated but we can't move to URL, because it doesn't
    // support server-relative URLs
    const queryString: ParsedUrlQuery = url.parse(request.url as string, true).query;
    const hasCode: boolean = queryString.code !== undefined;
    const hasError: boolean = queryString.error !== undefined;

    let body: string = "";
    if (hasCode === true) {
      body = '<script type="text/JavaScript">setTimeout(function(){ window.location = "https://pnp.github.io/cli-microsoft365/"; },10000);</script>';
      body += '<p><b>You have logged into CLI for Microsoft 365!</b></p>';
      body += '<p>You can close this window, or we will redirect you to the <a href="https://pnp.github.io/cli-microsoft365/">CLI for Microsoft 365</a> documentation in 10 seconds.</p>';

      this.resolve(<InteractiveAuthorizationCodeResponse>{
        code: queryString.code as string,
        redirectUri: this.generatedServerUrl
      });
    }

    if (hasError === true) {
      const errorMessage: ErrorResponse = {
        error: queryString.error as string,
        errorDescription: queryString.error_description as string
      }

      body = "<p>Oops! Azure Active Directory replied with an error message.</p>";
      body += `<p>${errorMessage.error}</p>`;
      if (errorMessage.errorDescription !== undefined) {
        body += `<p>${errorMessage.errorDescription}</p>`;
      }

      this.reject(errorMessage);
    }

    if (hasCode === false && hasError === false) {
      const errorMessage: ErrorResponse = {
        error: "invalid request",
        errorDescription: "An invalid request has been received by the HTTP server"
      }

      body = "<p>Oops! This is an invalid request.</p>";
      body += `<p>${errorMessage.error}</p>`;
      body += `<p>${errorMessage.errorDescription}</p>`;

      this.reject(errorMessage);
    }

    response.writeHead(200, { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'text/html' });
    response.write(`<html><head><title>CLI for Microsoft 365</title></head><body>${body}</body></html>`);
    response.end();

    this.httpServer.close();
  }
}

export default new AuthServer();