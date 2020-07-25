import commands from '../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./flow-export');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';
import * as fs from 'fs';

describe(commands.FLOW_EXPORT, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  const actualFilename = `20180916t000000zba9d7134cc81499e9884bf70642afac7_20180916042428.zip`
  const actualFileUrl = `https://bapfeblobprodml.blob.core.windows.net/20180916t000000zb5faa82a53cb4cd29f2a20fde7dbb785/${actualFilename}?sv=2017-04-17&sr=c&sig=AOp0fzKc0dLpY2yovI%2BSHJnQ92GxaMvbWgxyCX5Wwno%3D&se=2018-09-16T12%3A24%3A28Z&sp=rl`;
  const flowDisplayName = `Request manager approval for a Page`;
  const notFoundFlowId = '1c6ee23a-a835-44bc-a4f5-462b658efc12';
  const notFoundEnvironmentId = 'd87a7535-dd31-4437-bfe1-95340acd55c6';
  const foundFlowId = 'f2eb8b37-f624-4b22-9954-b5d0cbb28f8a';
  const foundEnvironmentId = 'cf409f12-a06f-426e-9955-20f5d7a31dd3';
  const nonZipFileFlowId = '694d21e4-49be-4e19-987b-074889e45c75';

  let postFakes = (opts: any) => {
    if ((opts.url as string).indexOf(notFoundEnvironmentId) > -1) {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": `Access to the environment 'Default-${notFoundEnvironmentId}' is denied.`
        }
      });
    }
    if (JSON.stringify(opts.body || {}).indexOf(notFoundFlowId) > -1) {
      return Promise.resolve({
        errors: [{
          "code": "ConnectionAuthorizationFailed",
          "message": `The caller with object id '${foundEnvironmentId}' does not have permission for connection '${notFoundFlowId}' under Api 'shared_logicflows'.`
        }]
      });
    }
    if ((opts.url as string).indexOf('/listPackageResources?api-version=2016-11-01') > -1) {
      return Promise.resolve(
        {
          "baseResourceIds": [`/providers/Microsoft.Flow/flows/${foundFlowId}`],
          "resources": { "L1BST1ZJREVSUy9NSUNST1NPRlQuRkxPVy9GTE9XUy9GMkVCOEIzNy1GNjI0LTRCMjItOTk1NC1CNUQwQ0JCMjhGOEI=": { "id": `/providers/Microsoft.Flow/flows/${foundFlowId}`, "name": `${foundFlowId}`, "type": "Microsoft.Flow/flows", "creationType": "Existing, New, Update", "details": { "displayName": flowDisplayName }, "configurableBy": "User", "hierarchy": "Root", "dependsOn": ["L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NIQVJFUE9JTlRPTkxJTkU=", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NIQVJFUE9JTlRPTkxJTkUvQ09OTkVDVElPTlMvU0hBUkVELVNIQVJFUE9JTlRPTkwtRjg0NTE4MDktREEwNi00RDQ3LTg3ODYtMTUxMjM4RDUwRTdB", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX09GRklDRTM2NVVTRVJT", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX09GRklDRTM2NVVTRVJTL0NPTk5FQ1RJT05TL1NIAAZAGFGH1FBSAHJKFS147VBDSxOUI5QjBELTFFQTUtNDhGOS1BQUM4LTgwRjkyQTFGRjE3OH==", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJADWNXX8321CGA3JIJDAkVEX0FQUFJPVkFMUw==", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX0FQUFJPVkFMUy9DT05ORUNUSU9OUy9TSEFSRUQtQVBQUk9WQUxTLUQ2Njc1AUUJNCSWDD1tNGNSAXZ1CNTY4LUFCRDc3MzMyOTMyMA==", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NFTkRNQUlM", "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NFTkRNQUlML0NPTk5FQ1RJT05TL1NIQVJFRC1TRU5ETUFJTC05NEUzODVCQi1CNUE3LTRBODgtOURFRC1FMEVFRDAzNTY1Njk="] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NIQVJFUE9JTlRPTkxJTkU=": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "name": "shared_sharepointonline", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1019.1195.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX09GRklDRTM2NVVTRVJT": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users", "name": "shared_office365users", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Microsoft 365 Users", "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1002.1175.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FAZFGGHDDCAAVEX0FQUFJPVkFMUw==": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals", "name": "shared_approvals", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Approvals", "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/Approvals3.svg" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NFTkRNQUlM": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail", "name": "shared_sendmail", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/officialicons/sendmail/icon_1.0.979.1161_83e4f20c-51d8-4c0c-a6f4-653249642047.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NIQVJFUE9JTlRPTkxJTkUvQ09OTkVDVElPTlMvU0hBUkVELVNIQVJFUE9JTlRPTkwtRjg0NTE4MDktREEwNi00RDQ3LTg3ODYtMTUxMjM4RDUwRTdB": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections/shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "name": "shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/sharepointonline.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NIQVJFUE9JTlRPTkxJTkU="] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX0FQUFJPVkFMUy9DT05ORUNUSU9OUy1AFFFVAGJXAAGHQUk9WQUxTLUQ2Njc1RUE5LUZDM0QtNDA4MS1CNTY4LUFCRDc3MzMyOTMyMZ==": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals/connections/shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "name": "shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Approvals", "iconUri": "https://connectorassets.blob.core.windows.net/assets/Approvals.svg" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hAZAASFCZ1DVHGVkFMUs=="] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NFTkRNQUlML0NPTk5FQ1RJT05TL1NIQVJFRC1TRU5ETUFJTC05NEUzODVCQi1CNUE3LTRBODgtOURFRC1FMEVFRDAzNTY1Njk=": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail/connections/shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "name": "shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/icons/sendmail.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX1NFTkRNQUlM"] }, "L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkAZVVDSSAFBRTM2NVVTRVJTL0NPTk5FQ1RJT05TL1NIQVJFRC1PRkZJQ0UzNjVVU0VSLUExOUI5QjBELTFFQTUtNDhGOS1BQUM4LTgwRjkyQTFGRjE3OB==": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users/connections/shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "name": "shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/office365users.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQSVMvU0hBUkVEX09GRklDRTM2NVVTRVJT"] } },
          "status": "Succeeded"
        }
      );
    }
    if ((opts.url as string).indexOf('/exportPackage?api-version=2016-11-01') > -1 && JSON.stringify(opts.body || {}).indexOf(nonZipFileFlowId) > -1) {
      return Promise.resolve(
        {
          "details": { "createdTime": "2018-09-16T04:24:28.365117Z", "packageTelemetryId": "448a7d93-7ce3-4e6a-88c9-57cf2479e62e" },
          "packageLink": { "value": `${actualFileUrl.replace('.zip', '.badextension')}` },
          "resources": { "43e3a371-ae70-455a-8050-4b14968ac474": { "id": `/providers/Microsoft.Flow/flows/${nonZipFileFlowId}`, "name": `${nonZipFileFlowId}`, "type": "Microsoft.Flow/flows", "status": "Succeeded", "creationType": "Existing, New, Update", "details": { "displayName": flowDisplayName }, "configurableBy": "User", "hierarchy": "Root", "dependsOn": ["0a6353d7-0770-447b-8d38-60230a1dc26d", "a6f57810-a099-4bf3-b51e-462afcea449e", "59eab504-a13a-40ed-b1f1-1decea0e1465", "1af3bf3f-97c9-4c45-b0fe-36613b9ff78c", "0e560c22-557c-432d-91a7-34f1562fc522", "e30ccca7-546e-4205-8e80-74f9f100b859", "94f5f489-8b4d-4e48-b50a-93514e16f921", "76995bea-58ce-4845-8298-1e29bf87e145"] }, "0a6353d7-0770-447b-8d38-60230a1dc26d": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "name": "shared_sharepointonline", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1019.1195.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "a6f57810-a099-4bf3-b51e-462afcea449e": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections/shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "name": "shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/sharepointonline.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["0a6353d7-0770-447b-8d38-60230a1dc26d"] }, "59eab504-a13a-40ed-b1f1-1decea0e1465": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users", "name": "shared_office365users", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Microsoft 365 Users", "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1002.1175.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "1af3bf3f-97c9-4c45-b0fe-36613b9ff78c": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users/connections/shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "name": "shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/office365users.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["59eab504-a13a-40ed-b1f1-1decea0e1465"] }, "0e560c22-557c-432d-91a7-34f1562fc522": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals", "name": "shared_approvals", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Approvals", "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/Approvals3.svg" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "e30ccca7-546e-4205-8e80-74f9f100b859": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals/connections/shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "name": "shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Approvals", "iconUri": "https://connectorassets.blob.core.windows.net/assets/Approvals.svg" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["0e560c22-557c-432d-91a7-34f1562fc522"] }, "94f5f489-8b4d-4e48-b50a-93514e16f921": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail", "name": "shared_sendmail", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/officialicons/sendmail/icon_1.0.979.1161_83e4f20c-51d8-4c0c-a6f4-653249642047.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "76995bea-58ce-4845-8298-1e29bf87e145": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail/connections/shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "name": "shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/icons/sendmail.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["94f5f489-8b4d-4e48-b50a-93514e16f921"] } },
          "status": "Succeeded"
        }
      );
    }
    if ((opts.url as string).indexOf('/exportPackage?api-version=2016-11-01') > -1) {
      return Promise.resolve(
        {
          "details": { "createdTime": "2018-09-16T04:24:28.365117Z", "packageTelemetryId": "448a7d93-7ce3-4e6a-88c9-57cf2479e62e" },
          "packageLink": { "value": `${actualFileUrl}` },
          "resources": { "43e3a371-ae70-455a-8050-4b14968ac474": { "id": `/providers/Microsoft.Flow/flows/${foundFlowId}`, "name": `${foundFlowId}`, "type": "Microsoft.Flow/flows", "status": "Succeeded", "creationType": "Existing, New, Update", "details": { "displayName": flowDisplayName }, "configurableBy": "User", "hierarchy": "Root", "dependsOn": ["0a6353d7-0770-447b-8d38-60230a1dc26d", "a6f57810-a099-4bf3-b51e-462afcea449e", "59eab504-a13a-40ed-b1f1-1decea0e1465", "1af3bf3f-97c9-4c45-b0fe-36613b9ff78c", "0e560c22-557c-432d-91a7-34f1562fc522", "e30ccca7-546e-4205-8e80-74f9f100b859", "94f5f489-8b4d-4e48-b50a-93514e16f921", "76995bea-58ce-4845-8298-1e29bf87e145"] }, "0a6353d7-0770-447b-8d38-60230a1dc26d": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "name": "shared_sharepointonline", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "SharePoint", "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1019.1195.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "a6f57810-a099-4bf3-b51e-462afcea449e": { "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections/shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "name": "shared-sharepointonl-f8451809-da06-4d47-8786-151238d50e7a", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/sharepointonline.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["0a6353d7-0770-447b-8d38-60230a1dc26d"] }, "59eab504-a13a-40ed-b1f1-1decea0e1465": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users", "name": "shared_office365users", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Microsoft 365 Users", "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1002.1175.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "1af3bf3f-97c9-4c45-b0fe-36613b9ff78c": { "id": "/providers/Microsoft.PowerApps/apis/shared_office365users/connections/shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "name": "shared-office365user-a19b9b0d-1ea5-48f9-aac8-80f92a1ff178", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "mark.powney@contoso.onmicrosoft.com", "iconUri": "https://az818438.vo.msecnd.net/icons/office365users.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["59eab504-a13a-40ed-b1f1-1decea0e1465"] }, "0e560c22-557c-432d-91a7-34f1562fc522": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals", "name": "shared_approvals", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Approvals", "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/Approvals3.svg" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "e30ccca7-546e-4205-8e80-74f9f100b859": { "id": "/providers/Microsoft.PowerApps/apis/shared_approvals/connections/shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "name": "shared-approvals-d6675ea9-fc3d-4081-b568-abd773329320", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Approvals", "iconUri": "https://connectorassets.blob.core.windows.net/assets/Approvals.svg" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["0e560c22-557c-432d-91a7-34f1562fc522"] }, "94f5f489-8b4d-4e48-b50a-93514e16f921": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail", "name": "shared_sendmail", "type": "Microsoft.PowerApps/apis", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/officialicons/sendmail/icon_1.0.979.1161_83e4f20c-51d8-4c0c-a6f4-653249642047.png" }, "configurableBy": "System", "hierarchy": "Child", "dependsOn": [] }, "76995bea-58ce-4845-8298-1e29bf87e145": { "id": "/providers/Microsoft.PowerApps/apis/shared_sendmail/connections/shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "name": "shared-sendmail-94e385bb-b5a7-4a88-9ded-e0eed0356569", "type": "Microsoft.PowerApps/apis/connections", "creationType": "Existing", "details": { "displayName": "Mail", "iconUri": "https://az818438.vo.msecnd.net/icons/sendmail.png" }, "configurableBy": "User", "hierarchy": "Child", "dependsOn": ["94f5f489-8b4d-4e48-b50a-93514e16f921"] } },
          "status": "Succeeded"
        }
      );
    }
    if ((opts.url as string).indexOf('/exportToARMTemplate?api-version=2016-11-01') > -1) {
      return Promise.resolve(
        {
        }
      );
    }
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf(notFoundEnvironmentId) > -1) {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": `Access to the environment 'Default-${notFoundEnvironmentId}' is denied.`
        }
      });
    }
    if ((opts.url as string).indexOf(notFoundFlowId) > -1) {
      return Promise.resolve({
        errors: [{
          "code": "ConnectionAuthorizationFailed",
          "message": `The caller with object id '${foundEnvironmentId}' does not have permission for connection '${notFoundFlowId}' under Api 'shared_logicflows'.`
        }]
      });
    }
    if (opts.url.match(/\/flows\/[^\?]+\?api-version\=2016-11-01/i)) {
      return Promise.resolve(
        {
          "id": `/providers/Microsoft.ProcessSimple/environments/Default-${foundEnvironmentId}/flows/${foundFlowId}`,
          "name": `${foundFlowId}`,
          "properties": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": flowDisplayName },
          "type": "Microsoft.ProcessSimple/environments/flows"
        }
      );
    }
    if (opts.url === actualFileUrl || opts.url === actualFileUrl.replace('.zip', '.badextension')) {
      return Promise.resolve('zipfilecontents');
    }

    return Promise.reject('Invalid request');
  }

  let writeFileSyncFake = () => { };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      fs.writeFileSync
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FLOW_EXPORT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('exports the specified flow (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`File saved to path './${actualFilename}'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exports flow to zip does not contain token', (done) => {
    const getRequestsStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip' } }, () => {
      try {
        assert.strictEqual(getRequestsStub.lastCall.args[0].headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exports the specified flow with a non zip file returned by the API (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, id: `${nonZipFileFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip', path: './output.zip', verbose: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`File saved to path './output.zip'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exports the specified flow in json format', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`./${flowDisplayName}.json`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exports the specified flow in json format (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: true, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`File saved to path './${flowDisplayName}.json'`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns ZIP file location when format specified as ZIP', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`./${actualFilename}`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('call is made without token when format specified as ZIP', (done) => {
    const getRequestsStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip' } }, () => {
      try {
        assert.strictEqual(getRequestsStub.lastCall.args[0].headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('nothing returned when path parameter is specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip', path: './output.zip' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('call is made without token when ZIP with specified path', (done) => {
    const getRequestsStub: sinon.SinonStub = sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);
    sinon.stub(fs, 'writeFileSync').callsFake(writeFileSyncFake);

    cmdInstance.action({ options: { debug: false, id: `${foundFlowId}`, environment: `Default-${foundEnvironmentId}`, format: 'zip', path: './output.zip' } }, () => {
      try {
        assert.strictEqual(getRequestsStub.lastCall.args[0].headers['x-anonymous'], true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({ options: { debug: false, environment: `Default-${notFoundEnvironmentId}`, id: `${foundFlowId}` } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-${notFoundEnvironmentId}' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({ options: { debug: false, environment: `Default-${foundEnvironmentId}`, id: notFoundFlowId } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The caller with object id '${foundEnvironmentId}' does not have permission for connection '${notFoundFlowId}' under Api 'shared_logicflows'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if format is specified as neither JSON nor ZIP', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'text' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if format is specified as JSON and packageCreatedBy parameter is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'json', packageCreatedBy: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if format is specified as JSON and packageDescription parameter is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'json', packageDescription: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if format is specified as JSON and packageDisplayName parameter is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'json', packageDisplayName: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if format is specified as JSON and packageSourceEnvironment parameter is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'json', packageSourceEnvironment: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, path: '/path/not/found.zip' } });
    Utils.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id and environment specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}` } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id and environment specified and format set to JSON', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: `Default-${foundEnvironmentId}`, id: `${foundFlowId}`, format: 'json' } });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying path', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--path') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying packageCreatedBy', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--packageCreatedBy') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying packageDescription', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--packageDescription') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying packageDisplayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--packageDisplayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying packageSourceEnvironment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--packageSourceEnvironment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});