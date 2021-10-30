import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001021_DEP_microsoft_sp_property_pane } from "./rules/FN001021_DEP_microsoft_sp_property_pane";
import { FN001022_DEP_office_ui_fabric_react } from "./rules/FN001022_DEP_office_ui_fabric_react";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base";
import { FN001024_DEP_microsoft_sp_diagnostics } from "./rules/FN001024_DEP_microsoft_sp_diagnostics";
import { FN001025_DEP_microsoft_sp_dynamic_data } from "./rules/FN001025_DEP_microsoft_sp_dynamic_data";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001028_DEP_microsoft_sp_list_subscription } from "./rules/FN001028_DEP_microsoft_sp_list_subscription";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN002009_DEVDEP_microsoft_sp_tslint_rules } from "./rules/FN002009_DEVDEP_microsoft_sp_tslint_rules";
import { FN002015_DEVDEP_types_react } from "./rules/FN002015_DEVDEP_types_react";
import { FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7 } from "./rules/FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7";
import { FN002018_DEVDEP_microsoft_rush_stack_compiler_3_9 } from "./rules/FN002018_DEVDEP_microsoft_rush_stack_compiler_3_9";
import { FN006004_CFG_PS_developer } from "./rules/FN006004_CFG_PS_developer";
import { FN007002_CFG_S_initialPage } from "./rules/FN007002_CFG_S_initialPage";
import { FN007003_CFG_S_api } from "./rules/FN007003_CFG_S_api";
import { FN008002_CFG_TSL_removeRule } from "./rules/FN008002_CFG_TSL_removeRule";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN012017_TSC_extends } from "./rules/FN012017_TSC_extends";
import { FN014007_CODE_launch_localWorkbench } from "./rules/FN014007_CODE_launch_localWorkbench";
import { FN015007_FILE_config_copy_assets_json } from "./rules/FN015007_FILE_config_copy_assets_json";
import { FN024001_NPMIGNORE_file } from "./rules/FN024001_NPMIGNORE_file";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.13.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.13.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.13.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.13.0'),
  new FN001008_DEP_react('16.13.1'),
  new FN001009_DEP_react_dom('16.13.1'),
  new FN001011_DEP_microsoft_sp_dialog('1.13.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.13.0'),
  new FN001013_DEP_microsoft_decorators('1.13.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.13.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.13.0'),
  new FN001022_DEP_office_ui_fabric_react('7.174.1'),
  new FN001023_DEP_microsoft_sp_component_base('1.13.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.13.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.13.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.13.0'),
  new FN001027_DEP_microsoft_sp_http('1.13.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.13.0'),
  new FN001029_DEP_microsoft_sp_loader('1.13.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.13.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.13.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.13.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.13.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.13.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('', false),
  new FN002009_DEVDEP_microsoft_sp_tslint_rules('1.13.0'),
  new FN002015_DEVDEP_types_react('16.9.51'),
  new FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7('', false),
  new FN002018_DEVDEP_microsoft_rush_stack_compiler_3_9('0.4.47'),
  new FN006004_CFG_PS_developer('1.13.0'),
  new FN007002_CFG_S_initialPage('https://enter-your-SharePoint-site/_layouts/workbench.aspx'),
  new FN007003_CFG_S_api(),
  new FN008002_CFG_TSL_removeRule('no-use-before-declare'),
  new FN010001_YORC_version('1.13.0'),
  new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-3.9/includes/tsconfig-web.json'),
  new FN014007_CODE_launch_localWorkbench(),
  new FN015007_FILE_config_copy_assets_json(false),
  new FN024001_NPMIGNORE_file()
];