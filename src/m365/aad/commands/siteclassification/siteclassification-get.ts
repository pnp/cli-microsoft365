import { GroupSetting, SettingValue } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { SiteClassificationSettings } from './SiteClassificationSettings';

class AadSiteClassificationGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_GET;
  }

  public get description(): string {
    return 'Gets site classification configuration';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: GroupSetting[] }>(requestOptions);
      if (res.value.length === 0) {
        throw 'Site classification is not enabled.';
      }

      const unifiedGroupSetting: GroupSetting[] = res.value.filter((directorySetting: GroupSetting): boolean => {
        return directorySetting.displayName === 'Group.Unified';
      });

      if (unifiedGroupSetting === null || unifiedGroupSetting.length === 0) {
        throw "Missing DirectorySettingTemplate for \"Group.Unified\"";
      }

      const siteClassificationsSettings: SiteClassificationSettings = new SiteClassificationSettings();

      // Get the classification list
      const classificationList: SettingValue[] = unifiedGroupSetting[0].values!.filter((directorySetting: SettingValue): boolean => {
        return directorySetting.name === 'ClassificationList';
      });

      siteClassificationsSettings.Classifications = [];
      if (classificationList !== null && classificationList.length > 0) {
        siteClassificationsSettings.Classifications = classificationList[0].value!.split(',');
      }

      // Get the UsageGuidelinesUrl
      const guidanceUrl: SettingValue[] = unifiedGroupSetting[0].values!.filter((directorySetting: SettingValue): boolean => {
        return directorySetting.name === 'UsageGuidelinesUrl';
      });

      siteClassificationsSettings.UsageGuidelinesUrl = "";
      if (guidanceUrl !== null && guidanceUrl.length > 0) {
        siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl[0]!.value!;
      }

      // Get the GuestUsageGuidelinesUrl
      const guestGuidanceUrl: SettingValue[] = unifiedGroupSetting[0].values!.filter((directorySetting: SettingValue): boolean => {
        return directorySetting.name === 'GuestUsageGuidelinesUrl';
      });

      siteClassificationsSettings.GuestUsageGuidelinesUrl = "";
      if (guestGuidanceUrl !== null && guestGuidanceUrl.length > 0) {
        siteClassificationsSettings.GuestUsageGuidelinesUrl = guestGuidanceUrl[0]!.value!;
      }

      // Get the DefaultClassification
      const defaultClassification: SettingValue[] = unifiedGroupSetting[0].values!.filter((directorySetting: SettingValue): boolean => {
        return directorySetting.name === 'DefaultClassification';
      });

      siteClassificationsSettings.DefaultClassification = "";
      if (defaultClassification !== null && defaultClassification.length > 0) {
        siteClassificationsSettings.DefaultClassification = defaultClassification[0].value!;
      }

      logger.log(JSON.parse(JSON.stringify(siteClassificationsSettings)));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadSiteClassificationGetCommand();