import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";
import { FN001027_DEP_microsoft_sp_http } from "./rules/FN001027_DEP_microsoft_sp_http";
import { FN001029_DEP_microsoft_sp_loader } from "./rules/FN001029_DEP_microsoft_sp_loader";
import { FN001030_DEP_microsoft_sp_module_interfaces } from "./rules/FN001030_DEP_microsoft_sp_module_interfaces";

module.exports = [
  new FN001004_DEP_microsoft_sp_webpart_base('1.1.1'),
  new FN001012_DEP_microsoft_sp_application_base('1.1.1'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('0.1.1'),
  new FN001027_DEP_microsoft_sp_http('1.1.1'),
  new FN001029_DEP_microsoft_sp_loader('1.1.1'),
  new FN001030_DEP_microsoft_sp_module_interfaces('1.1.1'),
  new FN010001_YORC_version('1.1.1')
];