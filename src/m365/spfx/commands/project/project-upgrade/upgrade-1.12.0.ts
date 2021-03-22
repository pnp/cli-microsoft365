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
import { FN002004_DEVDEP_gulp } from "./rules/FN002004_DEVDEP_gulp";
import { FN002005_DEVDEP_types_chai } from "./rules/FN002005_DEVDEP_types_chai";
import { FN002006_DEVDEP_types_mocha } from "./rules/FN002006_DEVDEP_types_mocha";
import { FN002009_DEVDEP_microsoft_sp_tslint_rules } from "./rules/FN002009_DEVDEP_microsoft_sp_tslint_rules";
import { FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3 } from "./rules/FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3";
import { FN002014_DEVDEP_types_es6_promise } from "./rules/FN002014_DEVDEP_types_es6_promise";
import { FN002015_DEVDEP_types_react } from "./rules/FN002015_DEVDEP_types_react";
import { FN002016_DEVDEP_types_react_dom } from "./rules/FN002016_DEVDEP_types_react_dom";
import { FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7 } from "./rules/FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN012013_TSC_exclude } from "./rules/FN012013_TSC_exclude";
import { FN012017_TSC_extends } from "./rules/FN012017_TSC_extends";
import { FN012018_TSC_lib_es2015_promise } from "./rules/FN012018_TSC_lib_es2015_promise";
import { FN012019_TSC_types_es6_promise } from "./rules/FN012019_TSC_types_es6_promise";
import { FN013002_GULP_serveTask } from "./rules/FN013002_GULP_serveTask";
import { FN015006_FILE_editorconfig } from "./rules/FN015006_FILE_editorconfig";
import { FN019002_TSL_extends } from "./rules/FN019002_TSL_extends";
import { FN021002_PKG_engines } from "./rules/FN021002_PKG_engines";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.12.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.12.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.12.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.12.0'),
  new FN001008_DEP_react('16.9.0'),
  new FN001009_DEP_react_dom('16.9.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.12.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.12.0'),
  new FN001013_DEP_microsoft_decorators('1.12.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.12.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.12.0'),
  new FN001022_DEP_office_ui_fabric_react('7.156.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.12.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.12.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.12.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.12.0'),
  new FN001027_DEP_microsoft_sp_http('1.12.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.12.0'),
  new FN001029_DEP_microsoft_sp_loader('1.12.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.12.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.12.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.12.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.12.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.12.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.12.0'),
  new FN002004_DEVDEP_gulp('4.0.2'),
  new FN002005_DEVDEP_types_chai('', false),
  new FN002006_DEVDEP_types_mocha('', false),
  new FN002009_DEVDEP_microsoft_sp_tslint_rules('1.12.0'),
  new FN002012_DEVDEP_microsoft_rush_stack_compiler_3_3('', false),
  new FN002017_DEVDEP_microsoft_rush_stack_compiler_3_7('0.2.3'),
  new FN002014_DEVDEP_types_es6_promise('', false),
  new FN002015_DEVDEP_types_react('16.9.36'),
  new FN002016_DEVDEP_types_react_dom('16.9.8'),
  new FN010001_YORC_version('1.12.0'),
  new FN012013_TSC_exclude([], false),
  new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-3.7/includes/tsconfig-web.json'),
  new FN012018_TSC_lib_es2015_promise(),
  new FN012019_TSC_types_es6_promise(false),
  new FN013002_GULP_serveTask(),
  new FN015006_FILE_editorconfig(false),
  new FN019002_TSL_extends('./node_modules/@microsoft/sp-tslint-rules/base-tslint.json'),
  new FN021002_PKG_engines(false)
];