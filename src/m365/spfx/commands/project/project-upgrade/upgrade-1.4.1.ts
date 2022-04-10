import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN002019_DEVDEP_spfx_fast_serve_helpers } from './rules/FN002019_DEVDEP_spfx_fast_serve_helpers';
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.4.1'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.4.1'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.4.1'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.4.1'),
  new FN001011_DEP_microsoft_sp_dialog('1.4.1'),
  new FN001012_DEP_microsoft_sp_application_base('1.4.1'),
  new FN001013_DEP_microsoft_decorators('1.4.1'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.4.1'),
  new FN001023_DEP_microsoft_sp_component_base('1.4.1'),
  new FN001026_DEP_microsoft_sp_extension_base('1.4.1'),
  new FN001027_DEP_microsoft_sp_http('1.4.1'),
  new FN001029_DEP_microsoft_sp_loader('1.4.1'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.4.1'),
  new FN001031_DEP_microsoft_sp_odata_types('1.4.1'),
  new FN001032_DEP_microsoft_sp_page_context('1.4.1'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.4.1'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.4.1'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.4.1'),
  new FN002019_DEVDEP_spfx_fast_serve_helpers('1.4.0'),
  new FN010001_YORC_version('1.4.1')
]; 