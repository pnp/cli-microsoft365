import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";

export default [
  new FN002001_DEVDEP_microsoft_sp_build_web({ packageVersion: '1.0.1' }),
  new FN010001_YORC_version('1.0.1')
];