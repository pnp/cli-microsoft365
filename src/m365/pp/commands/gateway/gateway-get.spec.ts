import * as assert from "assert";
import * as sinon from "sinon";
import appInsights from "../../../../appInsights";
import auth from "../../../../Auth";
import { Cli } from "../../../../cli/Cli";
import { CommandInfo } from "../../../../cli/CommandInfo";
import { Logger } from "../../../../cli/Logger";
import Command, { CommandError } from "../../../../Command";
import request from "../../../../request";
import { pid } from "../../../../utils/pid";
import { sinonUtil } from "../../../../utils/sinonUtil";
import commands from "../../commands";
const command: Command = require("./gateway-get");

describe(commands.GATEWAY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const gateway: any = {
    id: "1f69e798-5852-4fdd-ab01-33bb14b6e934",
    name: "My_Sample_Gateway",
    type: "Resource",
    publicKey: {
      exponent: "AQAB",
      modulus: "o6j2....cLk="
    }
  };

  before(() => {
    sinon.stub(auth, "restoreAuth").callsFake(() => Promise.resolve());
    sinon.stub(appInsights, "trackEvent").callsFake(() => {});
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogSpy = sinon.spy(logger, "log");
  });

  afterEach(() => {
    sinonUtil.restore([request.get]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth, 
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name.startsWith(commands.GATEWAY_GET), true);
  });

  it("has a description", () => {
    assert.notStrictEqual(command.description, null);
  });

  it("fails validation if the id is not valid", async () => {
    const actual = await command.validate(
      {
        options: {
          id: "3eb1a01b"
        }
      },
      commandInfo
    );

    assert.notStrictEqual(actual, true);
  });

  it("passes validation if the id is valid", async () => {
    const actual = await command.validate(
      {
        options: {
          id: "1f69e798-5852-4fdd-ab01-33bb14b6e934"
        }
      },
      commandInfo
    );

    assert.strictEqual(actual, true);
  });

  it("correctly handles error", async () => {
    sinon.stub(request, "get").callsFake(() => {
      return Promise.reject("An error has occurred");
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          id : 'testid'
        }
      }),
      new CommandError("An error has occurred")
    );
  });

  it("should get gateway information for the gateway by id", async () => {
    sinon.stub(request, "get").callsFake((opts) => {
      if (
        opts.url ===
        "https://api.powerbi.com" +
          `/v1.0/myorg/gateways/${encodeURIComponent(gateway.id)}`
      ) {
        return Promise.resolve(gateway);
      }
      return Promise.reject("Invalid request");
    });

    await command.action(logger, {
      options: {
        id: "1f69e798-5852-4fdd-ab01-33bb14b6e934"
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;

    assert.strictEqual(call.args[0].id, "1f69e798-5852-4fdd-ab01-33bb14b6e934");
    assert.strictEqual(call.args[0].name, "My_Sample_Gateway");
    assert(loggerLogSpy.calledWith(gateway));
  });
});
