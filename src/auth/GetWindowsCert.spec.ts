import { UseWindowsCerts } from "./GetWindowsCert";
import * as assert from "assert";

describe("Get personal certificate from Windows Store", () => {
  context("with thumbprint argument as string", () => {
    it("Returns Certificate from Windows Personal Store", (done) => {
      const thumbprint = "137BA5B7DDAB411BF7F50732CAA428270D335C51";
      const result = UseWindowsCerts(thumbprint);
      assert.strictEqual(typeof result, "string");
      done();
    });

    it("Thumbprint does not match", (done) => {
      const thumbprint = "thumbprint";
      const result = UseWindowsCerts(thumbprint);
      assert.strictEqual(result, "Certificate Not Found");
      done();
    });
  });
});