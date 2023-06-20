import { FN001001_DEP_microsoft_sp_core_library } from './rules/FN001001_DEP_microsoft_sp_core_library';
import { FN001002_DEP_microsoft_sp_lodash_subset } from './rules/FN001002_DEP_microsoft_sp_lodash_subset';
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from './rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core';
import { FN001004_DEP_microsoft_sp_webpart_base } from './rules/FN001004_DEP_microsoft_sp_webpart_base';
import { FN001011_DEP_microsoft_sp_dialog } from './rules/FN001011_DEP_microsoft_sp_dialog';
import { FN001012_DEP_microsoft_sp_application_base } from './rules/FN001012_DEP_microsoft_sp_application_base';
import { FN001013_DEP_microsoft_decorators } from './rules/FN001013_DEP_microsoft_decorators';
import { FN001014_DEP_microsoft_sp_listview_extensibility } from './rules/FN001014_DEP_microsoft_sp_listview_extensibility';
import { FN001021_DEP_microsoft_sp_property_pane } from './rules/FN001021_DEP_microsoft_sp_property_pane';
import { FN001023_DEP_microsoft_sp_component_base } from './rules/FN001023_DEP_microsoft_sp_component_base';
import { FN001024_DEP_microsoft_sp_diagnostics } from './rules/FN001024_DEP_microsoft_sp_diagnostics';
import { FN001025_DEP_microsoft_sp_dynamic_data } from './rules/FN001025_DEP_microsoft_sp_dynamic_data';
import { FN001026_DEP_microsoft_sp_extension_base } from './rules/FN001026_DEP_microsoft_sp_extension_base';
import { FN001027_DEP_microsoft_sp_http } from './rules/FN001027_DEP_microsoft_sp_http';
import { FN001028_DEP_microsoft_sp_list_subscription } from './rules/FN001028_DEP_microsoft_sp_list_subscription';
import { FN001029_DEP_microsoft_sp_loader } from './rules/FN001029_DEP_microsoft_sp_loader';
import { FN001030_DEP_microsoft_sp_module_interfaces } from './rules/FN001030_DEP_microsoft_sp_module_interfaces';
import { FN001031_DEP_microsoft_sp_odata_types } from './rules/FN001031_DEP_microsoft_sp_odata_types';
import { FN001032_DEP_microsoft_sp_page_context } from './rules/FN001032_DEP_microsoft_sp_page_context';
import { FN001034_DEP_microsoft_sp_adaptive_card_extension_base } from './rules/FN001034_DEP_microsoft_sp_adaptive_card_extension_base';
import { FN001035_DEP_fluentui_react } from './rules/FN001035_DEP_fluentui_react';
import { FN002001_DEVDEP_microsoft_sp_build_web } from './rules/FN002001_DEVDEP_microsoft_sp_build_web';
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from './rules/FN002002_DEVDEP_microsoft_sp_module_interfaces';
import { FN002020_DEVDEP_microsoft_rush_stack_compiler_4_5 } from './rules/FN002020_DEVDEP_microsoft_rush_stack_compiler_4_5';
import { FN002022_DEVDEP_microsoft_eslint_plugin_spfx } from './rules/FN002022_DEVDEP_microsoft_eslint_plugin_spfx';
import { FN002023_DEVDEP_microsoft_eslint_config_spfx } from './rules/FN002023_DEVDEP_microsoft_eslint_config_spfx';
import { FN002027_DEVDEP_fluentui_react } from './rules/FN002027_DEVDEP_fluentui_react';
import { FN010001_YORC_version } from './rules/FN010001_YORC_version';
import { FN014008_CODE_launch_hostedWorkbench_type } from './rules/FN014008_CODE_launch_hostedWorkbench_type';

module.exports = [
  new FN002027_DEVDEP_fluentui_react('', false),
  new FN001001_DEP_microsoft_sp_core_library('1.17.3'),
  new FN001023_DEP_microsoft_sp_component_base('1.17.3'),
  new FN001021_DEP_microsoft_sp_property_pane('1.17.3'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.17.3'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.17.3'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.17.3'),
  new FN001011_DEP_microsoft_sp_dialog('1.17.3'),
  new FN001012_DEP_microsoft_sp_application_base('1.17.3'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.17.3'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.17.3'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.17.3'),
  new FN001026_DEP_microsoft_sp_extension_base('1.17.3'),
  new FN001027_DEP_microsoft_sp_http('1.17.3'),
  new FN001028_DEP_microsoft_sp_list_subscription('1.17.3'),
  new FN001029_DEP_microsoft_sp_loader('1.17.3'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.17.3'),
  new FN001031_DEP_microsoft_sp_odata_types('1.17.3'),
  new FN001032_DEP_microsoft_sp_page_context('1.17.3'),
  new FN001013_DEP_microsoft_decorators('1.17.3'),
  new FN001034_DEP_microsoft_sp_adaptive_card_extension_base('1.17.3'),
  new FN001035_DEP_fluentui_react('7.199.1'),
  new FN002020_DEVDEP_microsoft_rush_stack_compiler_4_5('0.5.0'),
  new FN002022_DEVDEP_microsoft_eslint_plugin_spfx('1.17.3'),
  new FN002023_DEVDEP_microsoft_eslint_config_spfx('1.17.3'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.17.3'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.17.3'),
  new FN010001_YORC_version('1.17.3'),
  new FN014008_CODE_launch_hostedWorkbench_type('msedge')
];
