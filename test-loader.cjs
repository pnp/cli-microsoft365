require("source-map-support").install();
// https://sinonjs.org/how-to/stub-esm-default-export/
require = require("esm")(module, {
  cjs: true,
  mutableNamespace: true,
});
