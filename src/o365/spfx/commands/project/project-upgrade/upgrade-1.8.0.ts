import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN002009_DEVDEP_microsoft_sp_tslint_rules } from "./rules/FN002009_DEVDEP_microsoft_sp_tslint_rules";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom";
import { FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7 } from "./rules/FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7";
import { FN011011_MAN_webpart_supportedHosts } from "./rules/FN011011_MAN_webpart_supportedHosts";
import { FN012014_TSC_inlineSources } from "./rules/FN012014_TSC_inlineSources";
import { FN012015_TSC_strictNullChecks } from "./rules/FN012015_TSC_strictNullChecks";
import { FN012017_TSC_extends } from "./rules/FN012017_TSC_extends";
import { FN012016_TSC_noUnusedLocals } from "./rules/FN012016_TSC_noUnusedLocals";
import { FN001021_DEP_microsoft_sp_property_pane } from "./rules/FN001021_DEP_microsoft_sp_property_pane";
import { FN016004_TS_property_pane_property_import } from "./rules/FN016004_TS_property_pane_property_import";
import { FN018001_TEAMS_folder } from "./rules/FN018001_TEAMS_folder";
import { FN018003_TEAMS_tab20x20_png } from "./rules/FN018003_TEAMS_tab20x20_png";
import { FN018004_TEAMS_tab96x96_png } from "./rules/FN018004_TEAMS_tab96x96_png";
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

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.8.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.8.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.8.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.8.0'),
  new FN001021_DEP_microsoft_sp_property_pane('1.8.0'),
  new FN001008_DEP_react('16.7.0'),
  new FN001009_DEP_react_dom('16.7.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.8.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.8.0'),
  new FN001013_DEP_microsoft_decorators('1.8.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.8.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.8.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.8.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.8.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.8.0'),
  new FN001027_DEP_microsoft_sp_http('1.8.0'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.8.0'),
  new FN001029_DEP_microsoft_sp_loader('1.8.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.8.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.8.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.8.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.8.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.8.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.8.0'),
  new FN002009_DEVDEP_microsoft_sp_tslint_rules('1.8.0'),
  new FN002010_DEVDEP_microsoft_rush_stack_compiler_2_7('0.4.0'),
  new FN010001_YORC_version('1.8.0'),
  new FN011011_MAN_webpart_supportedHosts(true),
  new FN012014_TSC_inlineSources(false),
  new FN012015_TSC_strictNullChecks(false),
  new FN012016_TSC_noUnusedLocals(false),
  new FN012017_TSC_extends('./node_modules/@microsoft/rush-stack-compiler-2.7/includes/tsconfig-web.json'),
  new FN016004_TS_property_pane_property_import(),
  new FN018001_TEAMS_folder(),
  new FN018003_TEAMS_tab20x20_png(),
  new FN018004_TEAMS_tab96x96_png()
];