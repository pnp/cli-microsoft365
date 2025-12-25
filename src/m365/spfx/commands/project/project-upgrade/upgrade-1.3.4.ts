import { FN001001_DEP_microsoft_sp_core_library } from "./rules/FN001001_DEP_microsoft_sp_core_library.js";
import { FN001002_DEP_microsoft_sp_lodash_subset } from "./rules/FN001002_DEP_microsoft_sp_lodash_subset.js";
import { FN001003_DEP_microsoft_sp_office_ui_fabric_core } from "./rules/FN001003_DEP_microsoft_sp_office_ui_fabric_core.js";
import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base.js";
import { FN001011_DEP_microsoft_sp_dialog } from "./rules/FN001011_DEP_microsoft_sp_dialog.js";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base.js";
import { FN001013_DEP_microsoft_decorators } from "./rules/FN001013_DEP_microsoft_decorators.js";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility.js";
import { FN001023_DEP_microsoft_sp_component_base } from "./rules/FN001023_DEP_microsoft_sp_component_base.js";
import { FN001026_DEP_microsoft_sp_extension_base } from "./rules/FN001026_DEP_microsoft_sp_extension_base.js";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http.js";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader.js";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces.js";
import { FN001031_DEP_microsoft_sp_odata_types } from "./rules/FN001031_DEP_microsoft_sp_odata_types.js";
import { FN001032_DEP_microsoft_sp_page_context } from "./rules/FN001032_DEP_microsoft_sp_page_context.js";
import { FN002001_DEVDEP_microsoft_sp_build_web } from "./rules/FN002001_DEVDEP_microsoft_sp_build_web.js";
import { FN002002_DEVDEP_microsoft_sp_module_interfaces } from "./rules/FN002002_DEVDEP_microsoft_sp_module_interfaces.js";
import { FN002003_DEVDEP_microsoft_sp_webpart_workbench } from "./rules/FN002003_DEVDEP_microsoft_sp_webpart_workbench.js";
import { FN002007_DEVDEP_ajv } from "./rules/FN002007_DEVDEP_ajv.js";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version.js";
import { FN014002_CODE_extensions } from "./rules/FN014002_CODE_extensions.js";
import { FN014003_CODE_launch } from "./rules/FN014003_CODE_launch.js";

export default [
  new FN001001_DEP_microsoft_sp_core_library({ packageVersion: '1.3.4' }),
  new FN001002_DEP_microsoft_sp_lodash_subset({ packageVersion: '1.3.4' }),
  new FN001003_DEP_microsoft_sp_office_ui_fabric_core({ packageVersion: '1.3.4' }),
  new FN001004_DEP_microsoft_sp_webpart_base({ packageVersion: '1.3.4' }),
  new FN001011_DEP_microsoft_sp_dialog({ packageVersion: '1.3.4' }),
  new FN001012_DEP_microsoft_sp_application_base({ packageVersion: '1.3.4' }),
  new FN001013_DEP_microsoft_decorators({ packageVersion: '1.3.4' }),
  new FN001014_DEP_microsoft_sp_listview_extensibility({ packageVersion: '1.3.4' }),
  new FN001023_DEP_microsoft_sp_component_base({ packageVersion: '1.3.4' }),
  new FN001026_DEP_microsoft_sp_extension_base({ packageVersion: '1.3.4' }),
  new FN001027_DEP_microsoft_sp_http({ packageVersion: '1.3.4' }),
  new FN001029_DEP_microsoft_sp_loader({ packageVersion: '1.3.4' }),
  new FN001030_DEP_microsoft_sp_module_interfaces({ packageVersion: '1.3.4' }),
  new FN001031_DEP_microsoft_sp_odata_types({ packageVersion: '1.3.4' }),
  new FN001032_DEP_microsoft_sp_page_context({ packageVersion: '1.3.4' }),
  new FN002001_DEVDEP_microsoft_sp_build_web({ packageVersion: '1.3.4' }),
  new FN002002_DEVDEP_microsoft_sp_module_interfaces({ packageVersion: '1.3.4' }),
  new FN002003_DEVDEP_microsoft_sp_webpart_workbench({ packageVersion: '1.3.4' }),
  new FN002007_DEVDEP_ajv({ packageVersion: '5.2.2' }),
  new FN010001_YORC_version({ version: '1.3.4' }),
  new FN014002_CODE_extensions(),
  new FN014003_CODE_launch({ configuration: `{
  /**
    Install Chrome Debugger Extension for Visual Studio Code
    to debug your components with the Chrome browser:
    https://aka.ms/spfx-debugger-extensions
    */
  "version": "0.2.0",
  "configurations": [{
      "name": "Local workbench",
      "type": "chrome",
      "request": "launch",
      "url": "https://localhost:4321/temp/workbench.html",
      "webRoot": "\${workspaceRoot}",
      "sourceMaps": true,
      "sourceMapPathOverrides": {
        "webpack:///../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../../src/*": "\${webRoot}/src/*"
      },
      "runtimeArgs": [
        "--remote-debugging-port=9222"
      ]
    },
    {
      "name": "Hosted workbench",
      "type": "chrome",
      "request": "launch",
      "url": "https://enter-your-SharePoint-site/_layouts/workbench.aspx",
      "webRoot": "\${workspaceRoot}",
      "sourceMaps": true,
      "sourceMapPathOverrides": {
        "webpack:///../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../src/*": "\${webRoot}/src/*",
        "webpack:///../../../../../src/*": "\${webRoot}/src/*"
      },
      "runtimeArgs": [
        "--remote-debugging-port=9222",
        "-incognito"
      ]
    }
  ]
}` })
];