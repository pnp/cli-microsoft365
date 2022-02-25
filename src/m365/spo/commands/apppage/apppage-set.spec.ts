import * as assert from "assert";
import * as sinon from "sinon";
import appInsights from "../../../../appInsights";
import auth from "../../../../Auth";
import { Logger } from "../../../../cli";
import Command, { CommandError, CommandOption } from "../../../../Command";
import request from "../../../../request";
import { sinonUtil } from "../../../../utils";
import commands from "../../commands";
const command: Command = require("./apppage-set");

describe(commands.APPPAGE_SET, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, "restoreAuth").callsFake(() => Promise.resolve());
    sinon.stub(appInsights, "trackEvent").callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([request.post]);
  });

  after(() => {
    sinonUtil.restore([appInsights.trackEvent, auth.restoreAuth]);
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name.startsWith(commands.APPPAGE_SET), true);
  });

  it("has a description", () => {
    assert.notStrictEqual(command.description, null);
  });

  it("fails to update the single-part app page if request is rejected", done => {
    sinon.stub(request, "post").callsFake(opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1 &&
        opts.data.serverRelativeUrl.indexOf("failme")
      ) {
        return Promise.reject("Failed to update the single-part app page");
      }
      return Promise.reject("Invalid request");
    });
    command.action(logger, 
      {
        options: {
          debug: false,
          pageName: "failme",
          webUrl: "https://contoso.sharepoint.com/",
          webPartData: JSON.stringify({})
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(
            JSON.stringify(err),
            JSON.stringify(
              new CommandError(`Failed to update the single-part app page`)
            )
          );
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it("supports debug mode", () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach(o => {
      if (o.option === "--debug") {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying pageName", () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--pageName") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webUrl", () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webUrl") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webPartData", () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webPartData") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("fails validation if pageName not specified", () => {
    const actual = command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData not specified", () => {
    const actual = command.validate({
      options: {
        pageName: "Contoso.aspx",
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webUrl not specified", () => {
    const actual = command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        pageName: "page.aspx"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData is not a valid JSON string", () => {
    const actual = command.validate({
      options: {
        pageName: "Contoso.aspx",
        webUrl: "https://contoso",
        webPartData: "abc"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it("validation passes on all required options", () => {
    const actual = command.validate({
      options: {
        pageName: "Contoso.aspx",
        webPartData: "{}",
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.strictEqual(actual, true);
  });
});
