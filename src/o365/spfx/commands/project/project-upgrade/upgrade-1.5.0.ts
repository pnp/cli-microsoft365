import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN001007_DEP_types_webpack_env } from "./rules/FN001007_DEP_types_webpack_env";
import { FN001010_DEP_types_es6_promise } from "./rules/FN001010_DEP_types_es6_promise";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN002005_DEVDEP_types_chai } from "./rules/FN002005_DEVDEP_types_chai";
import { FN002006_DEVDEP_types_mocha } from "./rules/FN002006_DEVDEP_types_mocha";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators";
import { FN003001_CFG_schema } from "./rules/FN003001_CFG_schema";
import { FN004001_CFG_CA_schema } from "./rules/FN004001_CFG_CA_schema";
import { FN005001_CFG_DAS_schema } from "./rules/FN005001_CFG_DAS_schema";
import { FN006001_CFG_PS_schema } from "./rules/FN006001_CFG_PS_schema";
import { FN007001_CFG_S_schema } from "./rules/FN007001_CFG_S_schema";
import { FN008001_CFG_TSL_schema } from "./rules/FN008001_CFG_TSL_schema";
import { FN009001_CFG_WM_schema } from "./rules/FN009001_CFG_WM_schema";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN010002_YORC_isCreatingSolution } from "./rules/FN010002_YORC_isCreatingSolution";
import { FN010003_YORC_packageManager } from "./rules/FN010003_YORC_packageManager";
import { FN010004_YORC_componentType } from "./rules/FN010004_YORC_componentType";
import { FN011001_MAN_webpart_schema } from "./rules/FN011001_MAN_webpart_schema";
import { FN011002_MAN_applicationCustomizer_schema } from "./rules/FN011002_MAN_applicationCustomizer_schema";
import { FN011003_MAN_listViewCommandSet_schema } from "./rules/FN011003_MAN_listViewCommandSet_schema";
import { FN011004_MAN_fieldCustomizer_schema } from "./rules/FN011004_MAN_fieldCustomizer_schema";
import { FN012001_TSC_module } from "./rules/FN012001_TSC_module";
import { FN012002_TSC_moduleResolution } from "./rules/FN012002_TSC_moduleResolution";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base";
import { FN001024_DEP_microsoft_sp_diagnostics } from "./rules/FN001024_DEP_microsoft_sp_diagnostics";
import { FN001025_DEP_microsoft_sp_dynamic_data } from "./rules/FN001025_DEP_microsoft_sp_dynamic_data";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.5.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.5.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.5.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.5.0'),
  new FN001007_DEP_types_webpack_env('1.13.1'),
  new FN001010_DEP_types_es6_promise('0.0.33'),
  new FN001011_DEP_microsoft_sp_dialog('1.5.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.5.0'),
  new FN001013_DEP_microsoft_decorators('1.5.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.5.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.5.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.5.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.5.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.5.0'),
  new FN001027_DEP_microsoft_sp_http('1.5.0'),
  new FN001029_DEP_microsoft_sp_loader('1.5.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.5.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.5.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.5.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.5.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.5.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.5.0'),
  new FN002005_DEVDEP_types_chai('3.4.34'),
  new FN002006_DEVDEP_types_mocha('2.2.38'),
  new FN003001_CFG_schema('https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json'),
  new FN004001_CFG_CA_schema('https://developer.microsoft.com/json-schemas/spfx-build/copy-assets.schema.json'),
  new FN005001_CFG_DAS_schema('https://developer.microsoft.com/json-schemas/spfx-build/deploy-azure-storage.schema.json'),
  new FN006001_CFG_PS_schema('https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json'),
  new FN007001_CFG_S_schema('https://developer.microsoft.com/json-schemas/core-build/serve.schema.json'),
  new FN008001_CFG_TSL_schema('https://developer.microsoft.com/json-schemas/core-build/tslint.schema.json'),
  new FN009001_CFG_WM_schema('https://developer.microsoft.com/json-schemas/spfx-build/write-manifests.schema.json'),
  new FN010001_YORC_version('1.5.0'),
  new FN010002_YORC_isCreatingSolution(true),
  new FN010003_YORC_packageManager('npm'),
  new FN010004_YORC_componentType(),
  new FN011001_MAN_webpart_schema('https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json'),
  new FN011002_MAN_applicationCustomizer_schema('https://developer.microsoft.com/json-schemas/spfx/client-side-extension-manifest.schema.json'),
  new FN011003_MAN_listViewCommandSet_schema('https://developer.microsoft.com/json-schemas/spfx/command-set-extension-manifest.schema.json'),
  new FN011004_MAN_fieldCustomizer_schema('https://developer.microsoft.com/json-schemas/spfx/client-side-extension-manifest.schema.json'),
  new FN012001_TSC_module('esnext'),
  new FN012002_TSC_moduleResolution('node')
];