import { FN010201_CFG_TSL_preferConst } from "./rules/FN010201_CFG_TSL_preferConst";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN010203_DEP_knockout } from "./rules/FN010203_DEP_knockout";
import { FN010202_DEP_types_knockout } from "./rules/FN010202_DEP_types_knockout";

module.exports = [
  new FN010201_CFG_TSL_preferConst(),
  new FN010001_YORC_version('1.0.2'),
  new FN010203_DEP_knockout('3.4.0'),
  new FN010202_DEP_types_knockout('3.4.39')
];
