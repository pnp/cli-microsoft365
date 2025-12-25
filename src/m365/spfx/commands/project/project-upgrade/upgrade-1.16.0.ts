import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library.js";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset.js";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core.js";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base.js";
import { FN001008_DEP_react } from "./rules/FN001008_DEP_react.js";
import { FN001009_DEP_react_dom } from "./rules/FN001009_DEP_react_dom.js";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog.js";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base.js";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators.js";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility.js";
import { FN001021_DEP_microsoft_sp_property_pane } from "./rules/FN001021_DEP_microsoft_sp_property_pane.js";
import { FN001022_DEP_office_ui_fabric_react } from "./rules/FN001022_DEP_office_ui_fabric_react.js";
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
import { FN001034_DEP_microsoft_sp_adaptive_card_extension_base } from "./rules/FN001034_DEP_microsoft_sp_adaptive_card_extension_base.js";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN002015_DEVDEP_types_react } from "./rules/FN002015_DEVDEP_types_react.js";
import { FN002016_DEVDEP_types_react_dom } from "./rules/FN002016_DEVDEP_types_react_dom.js";
import { FN002019_DEVDEP_spfx_fast_serve_helpers } from "./rules/FN002019_DEVDEP_spfx_fast_serve_helpers.js";
import { FN002022_DEVDEP_microsoft_eslint_plugin_spfx } from "./rules/FN002022_DEVDEP_microsoft_eslint_plugin_spfx.js";
import { FN002023_DEVDEP_microsoft_eslint_config_spfx } from "./rules/FN002023_DEVDEP_microsoft_eslint_config_spfx.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN010008_YORC_nodeVersion } from "./rules/FN010008_YORC_nodeVersion.js";
import { FN010009_YORC_sdkVersions_microsoft_graph_client } from "./rules/FN010009_YORC_sdkVersions_microsoft_graph_client.js";
import { FN010010_YORC_sdkVersions_teams_js } from "./rules/FN010010_YORC_sdkVersions_teams_js.js";
import { FN021003_PKG_engines_node } from "./rules/FN021003_PKG_engines_node.js";
import { FN022001_SCSS_remove_fabric_react } from "./rules/FN022001_SCSS_remove_fabric_react.js";
import { FN022002_SCSS_add_fabric_react } from "./rules/FN022002_SCSS_add_fabric_react.js";

export default [
  new FN001001_DEP_microsoft_sp_core_library({ packageVersion: '1.16.0' }),
  new FN001002_DEP_microsoft_sp_lodash_subset({ packageVersion: '1.16.0' }),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core({ packageVersion: '1.16.0' }),
  new FN001004_DEP_microsoft_sp_webpart_base({ packageVersion: '1.16.0' }),
  new FN001008_DEP_react({ packageVersion: '17.0.1' }),
  new FN001009_DEP_react_dom({ packageVersion: '17.0.1' }),
  new FN001011_DEP_microsoft_sp_dialog({ packageVersion: '1.16.0' }),
  new FN001012_DEP_microsoft_sp_application_base({ packageVersion: '1.16.0' }),
  new FN001014_DEP_microsoft_sp_listview_extensibility({ packageVersion: '1.16.0' }),
  new FN001021_DEP_microsoft_sp_property_pane({ packageVersion: '1.16.0' }),
  new FN001022_DEP_office_ui_fabric_react({ packageVersion: '7.199.1' }),
  new FN001023_DEP_microsoft_sp_component_base({ packageVersion: '1.16.0' }),
  new FN001024_DEP_microsoft_sp_diagnostics({ packageVersion: '1.16.0' }),
  new FN001025_DEP_microsoft_sp_dynamic_data({ packageVersion: '1.16.0' }),
  new FN001026_DEP_microsoft_sp_extension_base({ packageVersion: '1.16.0' }),
  new FN001027_DEP_microsoft_sp_http({ packageVersion: '1.16.0' }),
  new FN001028_DEP_microsoft_sp_list_subscription({ packageVersion: '1.16.0' }),
  new FN001029_DEP_microsoft_sp_loader({ packageVersion: '1.16.0' }),
  new FN001030_DEP_microsoft_sp_module_interfaces({ packageVersion: '1.16.0' }),
  new FN001031_DEP_microsoft_sp_odata_types({ packageVersion: '1.16.0' }),
  new FN001032_DEP_microsoft_sp_page_context({ packageVersion: '1.16.0' }),
  new FN001013_DEP_microsoft_decorators({ packageVersion: '1.16.0' }),
  new FN001034_DEP_microsoft_sp_adaptive_card_extension_base({ packageVersion: '1.16.0' }),
  new FN002022_DEVDEP_microsoft_eslint_plugin_spfx({ packageVersion: '1.16.0' }),
  new FN002023_DEVDEP_microsoft_eslint_config_spfx({ packageVersion: '1.16.0' }),
  new FN002001_DEVDEP_microsoft_sp_build_web({ packageVersion: '1.16.0' }),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces({ packageVersion: '1.16.0' }),
  new FN002015_DEVDEP_types_react({ packageVersion: '17.0.45' }),
  new FN002016_DEVDEP_types_react_dom({ packageVersion: '17.0.17' }),
  new FN002019_DEVDEP_spfx_fast_serve_helpers({ packageVersion: '1.16.6' }),
  new FN010001_YORC_version({ version: '1.16.0' }),
  new FN010008_YORC_nodeVersion(),
  new FN010009_YORC_sdkVersions_microsoft_graph_client({ version: '3.0.2' }),
  new FN010010_YORC_sdkVersions_teams_js({ version: '2.4.1' }),
  new FN021003_PKG_engines_node('>=16.13.0 <17.0.0'),
  new FN022001_SCSS_remove_fabric_react({ oldValue: '~office-ui-fabric-react/dist/sass/References.scss' }),
  new FN022002_SCSS_add_fabric_react({ newValue: '~@fluentui/react/dist/sass/References.scss' })
];
