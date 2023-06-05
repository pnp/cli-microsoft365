import * as assert from "assert";
import * as sinon from "sinon";
import auth from "../../../../Auth";
import { Cli } from "../../../../cli/Cli";
import { CommandInfo } from "../../../../cli/CommandInfo";
import { Logger } from "../../../../cli/Logger";
import Command, { CommandError } from "../../../../Command";
import request from "../../../../request";
import { telemetry } from "../../../../telemetry";
import { formatting } from "../../../../utils/formatting";
import { pid } from "../../../../utils/pid";
import { session } from "../../../../utils/session";
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
    loggerLogSpy = sinon.spy(logger, "log");
  });

  afterEach(() => {
    sinonUtil.restore([request.get]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name, commands.GATEWAY_GET);
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
      throw "An error has occurred";
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          id: 'testid'
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
        `/v1.0/myorg/gateways/${formatting.encodeQueryParameter(gateway.id)}`
      ) {
        return gateway;
      }
      throw "Invalid request";
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
