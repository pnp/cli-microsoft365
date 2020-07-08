import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN003001_CFG_schema } from "./rules/FN003001_CFG_schema";
import { FN004001_CFG_CA_schema } from "./rules/FN004001_CFG_CA_schema";
import { FN005001_CFG_DAS_schema } from "./rules/FN005001_CFG_DAS_schema";
import { FN006001_CFG_PS_schema } from "./rules/FN006001_CFG_PS_schema";
import { FN007001_CFG_S_schema } from "./rules/FN007001_CFG_S_schema";
import { FN008001_CFG_TSL_schema } from "./rules/FN008001_CFG_TSL_schema";
import { FN008002_CFG_TSL_removeRule } from "./rules/FN008002_CFG_TSL_removeRule";
import { FN009001_CFG_WM_schema } from "./rules/FN009001_CFG_WM_schema";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN011001_MAN_webpart_schema } from "./rules/FN011001_MAN_webpart_schema";
import { FN011002_MAN_applicationCustomizer_schema } from "./rules/FN011002_MAN_applicationCustomizer_schema";
import { FN011003_MAN_listViewCommandSet_schema } from "./rules/FN011003_MAN_listViewCommandSet_schema";
import { FN011004_MAN_fieldCustomizer_schema } from "./rules/FN011004_MAN_fieldCustomizer_schema";
import { FN001005_DEP_types_react } from "./rules/FN001005_DEP_types_react";
import { FN003002_CFG_version } from "./rules/FN003002_CFG_version";
import { FN003003_CFG_bundles } from "./rules/FN003003_CFG_bundles";
import { FN003004_CFG_entries } from "./rules/FN003004_CFG_entries";
import { FN011006_MAN_listViewCommandSet_items } from "./rules/FN011006_MAN_listViewCommandSet_items";
import { FN011007_MAN_listViewCommandSet_removeCommands } from "./rules/FN011007_MAN_listViewCommandSet_removeCommands";
import { FN014004_CODE_settings_jsonSchemas_configJson_url } from "./rules/FN014004_CODE_settings_jsonSchemas_configJson_url";
import { FN003005_CFG_localizedResource_pathLib } from "./rules/FN003005_CFG_localizedResource_pathLib";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context";

module.exports = [
  new FN001001_DEP_microsoft_sp_core_library('1.2.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.2.0'),
  new FN001005_DEP_types_react('15.0.38'),
  new FN001012_DEP_microsoft_sp_application_base('1.2.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.2.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.2.0'),
  new FN001027_DEP_microsoft_sp_http('1.2.0'),
  new FN001029_DEP_microsoft_sp_loader('1.2.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.2.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.2.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.2.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.2.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.2.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.2.0'),
  new FN003001_CFG_schema('https://dev.office.com/json-schemas/spfx-build/config.2.0.schema.json'),
  new FN003002_CFG_version('2.0'),
  new FN003003_CFG_bundles(),
  new FN003004_CFG_entries(),
  new FN003005_CFG_localizedResource_pathLib(),
  new FN004001_CFG_CA_schema('https://dev.office.com/json-schemas/spfx-build/copy-assets.schema.json'),
  new FN005001_CFG_DAS_schema('https://dev.office.com/json-schemas/spfx-build/deploy-azure-storage.schema.json'),
  new FN006001_CFG_PS_schema('https://dev.office.com/json-schemas/spfx-build/package-solution.schema.json'),
  new FN007001_CFG_S_schema('https://dev.office.com/json-schemas/core-build/serve.schema.json'),
  new FN008002_CFG_TSL_removeRule('no-unused-imports'),
  new FN008001_CFG_TSL_schema('https://dev.office.com/json-schemas/core-build/tslint.schema.json'),
  new FN009001_CFG_WM_schema('https://dev.office.com/json-schemas/spfx-build/write-manifests.schema.json'),
  new FN010001_YORC_version('1.2.0'),
  new FN011001_MAN_webpart_schema('https://dev.office.com/json-schemas/spfx/client-side-web-part-manifest.schema.json'),
  new FN011002_MAN_applicationCustomizer_schema('https://dev.office.com/json-schemas/spfx/client-side-extension-manifest.schema.json'),
  new FN011003_MAN_listViewCommandSet_schema('https://dev.office.com/json-schemas/spfx/command-set-extension-manifest.schema.json'),
  new FN011004_MAN_fieldCustomizer_schema('https://dev.office.com/json-schemas/spfx/client-side-extension-manifest.schema.json'),
  new FN011006_MAN_listViewCommandSet_items(),
  new FN011007_MAN_listViewCommandSet_removeCommands(),
  new FN014004_CODE_settings_jsonSchemas_configJson_url('./node_modules/@microsoft/sp-build-core-tasks/lib/configJson/schemas/config-v1.schema.json')
];