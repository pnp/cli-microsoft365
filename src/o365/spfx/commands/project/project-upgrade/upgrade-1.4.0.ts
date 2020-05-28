import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN013001_GULP_msGridSassSuppression } from "./rules/FN013001_GULP_msGridSassSuppression";
import { FN014001_CODE_settings_jsonSchemas } from "./rules/FN014001_CODE_settings_jsonSchemas";
import { FN012003_TSC_skipLibCheck } from "./rules/FN012003_TSC_skipLibCheck";
import { FN006002_CFG_PS_includeClientSideAssets } from "./rules/FN006002_CFG_PS_includeClientSideAssets";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom";
import { FN001005_DEP_types_react } from "./rules/FN001005_DEP_types_react";
import { FN001006_DEP_types_react_dom } from "./rules/FN001006_DEP_types_react_dom";
import { FN012005_TSC_typeRoots_microsoft } from "./rules/FN012005_TSC_typeRoots_microsoft";
import { FN012004_TSC_typeRoots_types } from "./rules/FN012004_TSC_typeRoots_types";
import { FN012006_TSC_types_es6_collections } from "./rules/FN012006_TSC_types_es6_collections";
import { FN012007_TSC_lib_es5 } from "./rules/FN012007_TSC_lib_es5";
import { FN012008_TSC_lib_dom } from "./rules/FN012008_TSC_lib_dom";
import { FN012009_TSC_lib_es2015_collection } from "./rules/FN012009_TSC_lib_es2015_collection";
import { FN001015_DEP_types_react_addons_shallow_compare } from "./rules/FN001015_DEP_types_react_addons_shallow_compare";
import { FN001016_DEP_types_react_addons_update } from "./rules/FN001016_DEP_types_react_addons_update";
import { FN001017_DEP_types_react_addons_test_utils } from "./rules/FN001017_DEP_types_react_addons_update";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.4.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.4.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.4.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.4.0'),
  new FN001008_DEP_react('15.6.2'),
  new FN001009_DEP_react_dom('15.6.2'),
  new FN001005_DEP_types_react('15.6.6'),
  new FN001006_DEP_types_react_dom('15.5.6'),
  new FN001011_DEP_microsoft_sp_dialog('1.4.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.4.0'),
  new FN001013_DEP_microsoft_decorators('1.4.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.4.0'),
  new FN001015_DEP_types_react_addons_shallow_compare('', false),
  new FN001016_DEP_types_react_addons_update('', false),
  new FN001017_DEP_types_react_addons_test_utils('', false),
  new FN001023_DEP_microsoft_sp_component_base('1.4.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.4.0'),
  new FN001027_DEP_microsoft_sp_http('1.4.0'),
  new FN001029_DEP_microsoft_sp_loader('1.4.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.4.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.4.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.4.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.4.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.4.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.4.0'),
  new FN006002_CFG_PS_includeClientSideAssets(true),
  new FN010001_YORC_version('1.4.0'),
  new FN012003_TSC_skipLibCheck(true),
  new FN012004_TSC_typeRoots_types(),
  new FN012005_TSC_typeRoots_microsoft(),
  new FN012006_TSC_types_es6_collections(false),
  new FN012007_TSC_lib_es5(),
  new FN012008_TSC_lib_dom(),
  new FN012009_TSC_lib_es2015_collection(),
  new FN013001_GULP_msGridSassSuppression(),
  new FN014001_CODE_settings_jsonSchemas(false)
];