import { FN001004_DEP_microsoft_sp_webpart_base } from "./rules/FN001004_DEP_microsoft_sp_webpart_base";
import { FN001012_DEP_microsoft_sp_application_base } from "./rules/FN001012_DEP_microsoft_sp_application_base";
import { FN001014_DEP_microsoft_sp_listview_extensibility } from "./rules/FN001014_DEP_microsoft_sp_listview_extensibility";
import { FN010001_YORC_version } from "./rules/FN010001_YORC_version";

module.exports = [
  new FN001004_DEP_microsoft_sp_webpart_base('1.1.1'),
  new FN001012_DEP_microsoft_sp_application_base('1.1.1'),
  new FN001014_DEP_microsoft_sp_listview_extensibility('0.1.1'),
  new FN010001_YORC_version('1.1.1'),
];