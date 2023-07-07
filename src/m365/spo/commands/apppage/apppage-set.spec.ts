import * as assert from "assert";
import * as sinon from "sinon";
import { telemetry } from "../../../../telemetry";
import auth from "../../../../Auth";
import { Cli } from "../../../../cli/Cli";
import { CommandInfo } from "../../../../cli/CommandInfo";
import { Logger } from "../../../../cli/Logger";
import Command, { CommandError } from "../../../../Command";
import request from "../../../../request";
import { sinonUtil } from "../../../../utils/sinonUtil";
import commands from "../../commands";
import { session } from "../../../../utils/session";
import { pid } from "../../../../utils/pid";
const command: Command = require("./apppage-set");

describe(commands.APPPAGE_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name, commands.APPPAGE_SET);
  });

  it("has a description", () => {
    assert.notStrictEqual(command.description, null);
  });

  it("fails to update the single-part app page if request is rejected", async () => {
    sinon.stub(request, "post").callsFake(async opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1 &&
        opts.data.serverRelativeUrl.indexOf("failme")
      ) {
        throw "Failed to update the single-part app page";
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger,
      {
        options: {
          name: "failme",
          webUrl: "https://contoso.sharepoint.com/",
          webPartData: JSON.stringify({})
        }
      }), new CommandError(`Failed to update the single-part app page`));
  });

  it("Update the single-part app pag", async () => {
    sinon.stub(request, "post").callsFake(async opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1
      ) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger,
      {
        options: {
          pageName: "demo",
          webUrl: "https://contoso.sharepoint.com/teams/sales",
          webPartData: JSON.stringify({})
        }
      });
  });

  it("supports specifying name", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--name") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webUrl", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webUrl") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webPartData", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webPartData") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("fails validation if name not specified", async () => {
    const actual = await command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData not specified", async () => {
    const actual = await command.validate({
      options: {
        name: "Contoso.aspx",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webUrl not specified", async () => {
    const actual = await command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        name: "page.aspx"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData is not a valid JSON string", async () => {
    const actual = await command.validate({
      options: {
        name: "Contoso.aspx",
        webUrl: "https://contoso",
        webPartData: "abc"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("validation passes on all required options", async () => {
    const actual = await command.validate({
      options: {
        name: "Contoso.aspx",
        webPartData: "{}",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
