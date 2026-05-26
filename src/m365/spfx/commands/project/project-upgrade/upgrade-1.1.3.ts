import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN011008_MAN_requiresCustomScript } from "./rules/FN011008_MAN_requiresCustomScript.js";
import { FN011009_MAN_webpart_safeWithCustomScriptDisabled } from "./rules/FN011009_MAN_webpart_safeWithCustomScriptDisabled.js";

export default [
  new FN002002_DEVDEP_microsoft_sp_module_interfaces({ packageVersion: '1.1.1' }),
  new FN010001_YORC_version({ version: '1.1.3' }),
  new FN011008_MAN_requiresCustomScript(),
  new FN011009_MAN_webpart_safeWithCustomScriptDisabled({ add: false })
];