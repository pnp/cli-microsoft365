import commands from "../../commands";
import Command, { CommandOption, CommandValidate, CommandError } from "../../../../Command";
import * as sinon from "sinon";
import appInsights from "../../../../appInsights";
import auth from "../../../../Auth";
const command: Command = require("./apppage-set");
import * as assert from "assert";
import request from "../../../../request";
import Utils from "../../../../Utils";

describe(commands.APPPAGE_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, "restoreAuth").callsFake(() => Promise.resolve());
    sinon.stub(appInsights, "trackEvent").callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require("../../../../vorpal-init");
    log = [];
    cmdInstance = {
      commandWrapper: { command: command.name },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([vorpal.find, request.post]);
  });

  after(() => {
    Utils.restore([appInsights.trackEvent, auth.restoreAuth]);
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.equal(command.name.startsWith(commands.APPPAGE_SET), true);
  });

  it("has a description", () => {
    assert.notEqual(command.description, null);
  });

  it("updates the single-part app page", done => {
    sinon.stub(request, "post").callsFake(opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1 &&
        opts.body.webPartDataAsJson === "{}"
      ) {
        return Promise.resolve("Done");
      }
      return Promise.reject("Invalid request");
    });
    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: "test-single",
          webUrl: "https://contoso.sharepoint.com/",
          webPartData: JSON.stringify({})
        }
      },
      (err?: any) => {
        if (err) {
          done(err);
          return;
        }
        try {
          assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it("fails to update the single-part app page if request is rejected", done => {
    sinon.stub(request, "post").callsFake(opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1 &&
        opts.body.serverRelativeUrl.indexOf("failme")
      ) {
        return Promise.reject("Failed to update the single-part app page");
      }
      return Promise.reject("Invalid request");
    });
    cmdInstance.action(
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
          assert.equal(
            JSON.stringify(err),
            JSON.stringify(
              new CommandError(`Failed to update the single-part app page`)
            )
          );
          done();
        } catch (e) {
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
    const actual = (command.validate() as CommandValidate)({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.notEqual(actual, true);
  });

  it("fails validation if webPartData not specified", () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: "Contoso.aspx",
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.notEqual(actual, true);
  });

  it("fails validation if webUrl not specified", () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        pageName: "page.aspx"
      }
    });
    assert.notEqual(actual, true);
  });

  it("fails validation if webPartData is not a valid JSON string", () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: "Contoso.aspx",
        webUrl: "https://contoso",
        webPartData: "abc"
      }
    });
    assert.notEqual(actual, true);
  });

  it("validation passes on all required options", () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: "Contoso.aspx",
        webPartData: "{}",
        webUrl: "https://contoso.sharepoint.com"
      }
    });
    assert.equal(actual, true);
  });

  it("has help referring to the right command", () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, "find").callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.APPPAGE_SET));
  });

  it("has help with examples", () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, "find").callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf("Examples:") > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});
