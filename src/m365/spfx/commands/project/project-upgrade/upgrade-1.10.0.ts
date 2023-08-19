import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library.js";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset.js";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core.js";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base.js";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog.js";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base.js";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators.js";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility.js";
import { FN001021_DEP_microsoft_sp_property_pane } from "./rules/FN001021_DEP_microsoft_sp_property_pane.js";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base.js";
import { FN001024_DEP_microsoft_sp_diagnostics } from "./rules/FN001024_DEP_microsoft_sp_diagnostics.js";
import { FN001025_DEP_microsoft_sp_dynamic_data } from "./rules/FN001025_DEP_microsoft_sp_dynamic_data.js";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base.js";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http.js";
import { FN001028_DEP_microsoft_sp_list_subscription } from "./rules/FN001028_DEP_microsoft_sp_list_subscription.js";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader.js";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces.js";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types.js";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context.js";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench.js";
import { FN002009_DEVDEP_microsoft_sp_tslint_rules } from "./rules/FN002009_DEVDEP_microsoft_sp_tslint_rules.js";
import { FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9 } from "./rules/FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9.js";
import { FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3 } from "./rules/FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN012017_TSC_extends } from "./rules/FN012017_TSC_extends.js";
import { FN016004_TS_property_pane_property_import } from "./rules/FN016004_TS_property_pane_property_import.js";

export default [
  new FN001001_DEP_microsoft_sp_core_library('1.10.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.10.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.10.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.10.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.10.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.10.0'),
  new FN001013_DEP_microsoft_decorators('1.10.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.10.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.10.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.10.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.10.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.10.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.10.0'),
  new FN001027_DEP_microsoft_sp_http('1.10.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.10.0'),
  new FN001029_DEP_microsoft_sp_loader('1.10.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.10.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.10.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.10.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.10.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.10.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.10.0'),
  new FN002009_DEVDEP_microsoft_sp_tslint_rules('1.10.0'),
  new FN002011_DEVDEP_microsoft_rush_stack_compiler_2_9('0.7.16', true),
  new FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3('0.3.5'),
  new FN010001_YORC_version('1.10.0'),
  new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-3.3/includes/tsconfig-web.json'),
  new FN016004_TS_property_pane_property_import()
];