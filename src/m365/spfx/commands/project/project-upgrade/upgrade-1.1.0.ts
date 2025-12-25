import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library.js";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base.js";
import { FN001005_DEP_types_react } from "./rules/FN001005_DEP_types_react.js";
import { FN001006_DEP_types_react_dom } from "./rules/FN001006_DEP_types_react_dom.js";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react.js";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom.js";
import { FN001015_DEP_types_react_addons_shallow_compare } from "./rules/FN001015_DEP_types_react_addons_shallow_compare.js";
import { FN001016_DEP_types_react_addons_update } from "./rules/FN001016_DEP_types_react_addons_update.js";
import { FN001017_DEP_types_react_addons_test_utils } from "./rules/FN001017_DEP_types_react_addons_update.js";
import { FN001018_DEP_microsoft_sp_client_base } from "./rules/FN001018_DEP_microsoft_sp_client_base.js";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base.js";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http.js";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader.js";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces.js";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types.js";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context.js";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN010005_YORC_environment } from "./rules/FN010005_YORC_environment.js";
import { FN010006_YORC_framework } from "./rules/FN010006_YORC_framework.js";
import { FN011009_MAN_webpart_safeWithCustomScriptDisabled } from "./rules/FN011009_MAN_webpart_safeWithCustomScriptDisabled.js";
import { FN011010_MAN_webpart_version } from "./rules/FN011010_MAN_webpart_version.js";
import { FN012010_TSC_experimentalDecorators } from "./rules/FN012010_TSC_experimentalDecorators.js";
import { FN014005_CODE_settingsfile } from "./rules/FN014005_CODE_settingsfile.js";
import { FN015001_FILE_typings_tsd_d_ts } from "./rules/FN015001_FILE_typings_tsd_d_ts.js";
import { FN015002_FILE_typings__ms_odsp_d_ts } from "./rules/FN015002_FILE_typings__ms_odsp_d_ts.js";

export default [
  new FN001001_DEP_microsoft_sp_core_library({ packageVersion: '1.1.0' }),
  new FN001004_DEP_microsoft_sp_webpart_base({ packageVersion: '1.1.0' }),
  new FN001005_DEP_types_react({ packageVersion: '0.14.46' }),
  new FN001006_DEP_types_react_dom({ packageVersion: '0.14.18' }),
  new FN001008_DEP_react({ packageVersion: '15.4.2' }),
  new FN001009_DEP_react_dom({ packageVersion: '15.4.2' }),
  new FN001015_DEP_types_react_addons_shallow_compare('0.14.17', true),
  new FN001016_DEP_types_react_addons_update('0.14.14', true),
  new FN001017_DEP_types_react_addons_test_utils('0.14.15', true),
  new FN001018_DEP_microsoft_sp_client_base('', false),
  new FN001023_DEP_microsoft_sp_component_base({ packageVersion: '1.1.0' }),
  new FN001027_DEP_microsoft_sp_http({ packageVersion: '1.1.0' }),
  new FN001029_DEP_microsoft_sp_loader({ packageVersion: '1.1.0' }),
  new FN001030_DEP_microsoft_sp_module_interfaces({ packageVersion: '1.1.0' }),
  new FN001031_DEP_microsoft_sp_odata_types({ packageVersion: '1.1.0' }),
  new FN001032_DEP_microsoft_sp_page_context({ packageVersion: '1.1.0' }),
  new FN002001_DEVDEP_microsoft_sp_build_web({ packageVersion: '1.1.0' }),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces({ packageVersion: '1.1.0' }),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench({ packageVersion: '1.1.0' }),
  new FN010001_YORC_version('1.1.0'),
  new FN010005_YORC_environment('spo'),
  new FN010006_YORC_framework('', false),
  new FN011009_MAN_webpart_safeWithCustomScriptDisabled(true),
  new FN011010_MAN_webpart_version(),
  new FN012010_TSC_experimentalDecorators(),
  new FN014005_CODE_settingsfile(),
  new FN015001_FILE_typings_tsd_d_ts({ add: false }),
  new FN015002_FILE_typings__ms_odsp_d_ts({ add: false })
];