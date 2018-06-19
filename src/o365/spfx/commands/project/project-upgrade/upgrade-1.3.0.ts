import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN006003_CFG_PS_skipFeatureDeployment } from "./rules/FN006003_CFG_PS_skipFeatureDeployment";
import { FN011005_MAN_webpart_defaultGroup } from "./rules/FN011005_MAN_webpart_defaultGroup";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.3.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.3.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.3.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.3.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.3.0'),
  new FN001013_DEP_microsoft_decorators('1.3.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.3.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.3.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.3.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.3.0'),
  new FN006003_CFG_PS_skipFeatureDeployment('string'),
  new FN010001_YORC_version('1.3.0'),
  new FN011005_MAN_webpart_defaultGroup('Under Development', 'Other')
];