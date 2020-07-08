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
import { FN002008_DEVDEP_tslint_microsoft_contrib } from "./rules/FN002008_DEVDEP_tslint_microsoft_contrib";
import { FN012011_TSC_outDir } from "./rules/FN012011_TSC_outDir";
import { FN012012_TSC_include } from "./rules/FN012012_TSC_include";
import { FN012013_TSC_exclude } from "./rules/FN012013_TSC_exclude";
import { FN015003_FILE_tslint_json } from "./rules/FN015003_FILE_tslint_json";
import { FN015004_FILE_config_tslint_json } from "./rules/FN015004_FILE_config_tslint_json";
import { FN015005_FILE_src_index_ts } from "./rules/FN015005_FILE_src_index_ts";
import { FN016001_TS_msgraphclient_packageName } from "./rules/FN016001_TS_msgraphclient_packageName";
import { FN016002_TS_msgraphclient_instance } from "./rules/FN016002_TS_msgraphclient_instance";
import { FN016003_TS_aadhttpclient_instance } from "./rules/FN016003_TS_aadhttpclient_instance";
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
  new FN001001_DEP_microsoft_sp_core_library('1.6.0'),
  new FN001002_DEP_microsoft_sp_lodash_subset('1.6.0'),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core('1.6.0'),
  new FN001004_DEP_microsoft_sp_webpart_base('1.6.0'),
  new FN001011_DEP_microsoft_sp_dialog('1.6.0'),
  new FN001012_DEP_microsoft_sp_application_base('1.6.0'),
  new FN001013_DEP_microsoft_decorators('1.6.0'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('1.6.0'),
  new FN001023_DEP_microsoft_sp_component_base('1.6.0'),
  new FN001024_DEP_microsoft_sp_diagnostics('1.6.0'),
  new FN001025_DEP_microsoft_sp_dynamic_data('1.6.0'),
  new FN001026_DEP_microsoft_sp_extension_base('1.6.0'),
  new FN001027_DEP_microsoft_sp_http('1.6.0'),
  new FN001029_DEP_microsoft_sp_loader('1.6.0'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.6.0'),
  new FN001031_DEP_microsoft_sp_odata_types('1.6.0'),
  new FN001032_DEP_microsoft_sp_page_context('1.6.0'),
  new FN002001_DEVDEP_microsoft_sp_build_web('1.6.0'),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces('1.6.0'),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench('1.6.0'),
  new FN002008_DEVDEP_tslint_microsoft_contrib('5.0.0'),
  new FN010001_YORC_version('1.6.0'),
  new FN012011_TSC_outDir('lib'),
  new FN012012_TSC_include([
    'src/**/*.ts'
  ]),
  new FN012013_TSC_exclude([
    'node_modules',
    'lib'
  ]),
  new FN015003_FILE_tslint_json(true, `{
  "rulesDirectory": [
    "tslint-microsoft-contrib"
  ],
  "rules": {
    "class-name": false,
    "export-name": false,
    "forin": false,
    "label-position": false,
    "member-access": true,
    "no-arg": false,
    "no-console": false,
    "no-construct": false,
    "no-duplicate-variable": true,
    "no-eval": false,
    "no-function-expression": true,
    "no-internal-module": true,
    "no-shadowed-variable": true,
    "no-switch-case-fall-through": true,
    "no-unnecessary-semicolons": true,
    "no-unused-expression": true,
    "no-use-before-declare": true,
    "no-with-statement": true,
    "semicolon": true,
    "trailing-comma": false,
    "typedef": false,
    "typedef-whitespace": false,
    "use-named-parameter": true,
    "variable-name": false,
    "whitespace": false
  }
}`),
  new FN015004_FILE_config_tslint_json(false),
  new FN015005_FILE_src_index_ts(true, `// A file is required to be in the root of the /src directory by the TypeScript compiler
`),
  new FN016001_TS_msgraphclient_packageName('@microsoft/sp-http'),
  new FN016002_TS_msgraphclient_instance(),
  new FN016003_TS_aadhttpclient_instance()
];