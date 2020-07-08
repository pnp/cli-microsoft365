import { FN008003_CFG_TSL_preferConst } from "./rules/FN008003_CFG_TSL_preferConst";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN001019_DEP_knockout } from "./rules/FN001019_DEP_knockout";
import { FN001020_DEP_types_knockout } from "./rules/FN001020_DEP_types_knockout";

module.exports = [
  new FN001019_DEP_knockout('3.4.0'),
  new FN001020_DEP_types_knockout('3.4.39'),
  new FN008003_CFG_TSL_preferConst(),
  new FN010001_YORC_version('1.0.2')
];
