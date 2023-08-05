import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library.js";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset.js";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core.js";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base.js";
import { FN001005_DEP_types_react } from "./rules/FN001005_DEP_types_react.js";
import { FN001006_DEP_types_react_dom } from "./rules/FN001006_DEP_types_react_dom.js";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react.js";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom.js";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog.js";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base.js";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators.js";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility.js";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base.js";
import { FN001024_DEP_microsoft_sp_diagnostics } from "./rules/FN001024_DEP_microsoft_sp_diagnostics.js";
import { FN001025_DEP_microsoft_sp_dynamic_data } from "./rules/FN001025_DEP_microsoft_sp_dynamic_data.js";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base.js";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http.js";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader.js";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces.js";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types.js";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context.js";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench.js";
import { FN002008_DEVDEP_tslint_microsoft_contrib } from "./rules/FN002008_DEVDEP_tslint_microsoft_contrib.js";
import { FN002009_DEVDEP_microsoft_sp_tslint_rules } from "./rules/FN002009_DEVDEP_microsoft_sp_tslint_rules.js";
import { FN006003_CFG_PS_isDomainIsolated } from "./rules/FN006003_CFG_PS_isDomainIsolated.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN010007_YORC_isDomainIsolated } from "./rules/FN010007_YORC_isDomainIsolated.js";
import { FN018001_TEAMS_folder } from "./rules/FN018001_TEAMS_folder.js";
import { FN018002_TEAMS_manifest } from "./rules/FN018002_TEAMS_manifest.js";
import { FN018003_TEAMS_tab20x20_png } from "./rules/FN018003_TEAMS_tab20x20_png.js";
import { FN018004_TEAMS_tab96x96_png } from "./rules/FN018004_TEAMS_tab96x96_png.js";
import { FN019001_TSL_rulesDirectory } from "./rules/FN019001_TSL_rulesDirectory.js";
import { FN019002_TSL_extends } from "./rules/FN019002_TSL_extends.js";

export default [
  new FN001001_DEP_microsoft_sp_core_library('1.7.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.7.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.7.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.7.0'),
  new FN001005_DEP_types_react('16.4.2'),
  new FN001006_DEP_types_react_dom('16.0.5'),
  new FN001008_DEP_react('16.3.2'),
  new FN001009_DEP_react_dom('16.3.2'),
  new FN001011_DEP_microsoft_sp_dialog('1.7.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.7.0'),
  new FN001013_DEP_microsoft_decorators('1.7.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.7.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.7.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.7.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.7.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.7.0'),
  new FN001027_DEP_microsoft_sp_http('1.7.0'),
  new FN001029_DEP_microsoft_sp_loader('1.7.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.7.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.7.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.7.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.7.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.7.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.7.0'),
  new FN002008_DEVDEP_tslint_microsoft_contrib('', false),
  new FN002009_DEVDEP_microsoft_sp_tslint_rules('1.7.0'),
  new FN006003_CFG_PS_isDomainIsolated(false),
  new FN010001_YORC_version('1.7.0'),
  new FN010007_YORC_isDomainIsolated(false),
  new FN018001_TEAMS_folder(),
  new FN018002_TEAMS_manifest(),
  new FN018003_TEAMS_tab20x20_png('tab20x20.png'),
  new FN018004_TEAMS_tab96x96_png('tab96x96.png'),
  new FN019001_TSL_rulesDirectory(),
  new FN019002_TSL_extends('@microsoft/sp-tslint-rules/base-tslint.json')
];