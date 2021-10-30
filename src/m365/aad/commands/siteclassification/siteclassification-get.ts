import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { DirectorySetting } from './DirectorySetting';
import { DirectorySettingTemplatesRsp } from './DirectorySettingTemplatesRsp';
import { DirectorySettingValue } from './DirectorySettingValue';
import { SiteClassificationSettings } from './SiteClassificationSettings';

interface CommandArgs {
  options: GlobalOptions;
}

class AadSiteClassificationGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_GET;
  }

  public get description(): string {
    return 'Gets site classification configuration';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/groupSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<DirectorySettingTemplatesRsp>(requestOptions)
      .then((res: DirectorySettingTemplatesRsp): void => {
        if (res.value.length === 0) {
          cb(new CommandError('Site classification is not enabled.'));
          return;
        }

        const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
          return directorySetting.displayName === 'Group.Unified';
        });

        if (unifiedGroupSetting === null || unifiedGroupSetting.length === 0) {
          cb(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
          return;
        }

        const siteClassificationsSettings: SiteClassificationSettings = new SiteClassificationSettings();

        // Get the classification list
        const classificationList: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
          return directorySetting.name === 'ClassificationList';
        });

        siteClassificationsSettings.Classifications = [];
        if (classificationList !== null && classificationList.length > 0) {
          siteClassificationsSettings.Classifications = classificationList[0].value.split(',');
        }

        // Get the UsageGuidelinesUrl
        const guidanceUrl: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
          return directorySetting.name === 'UsageGuidelinesUrl';
        });

        siteClassificationsSettings.UsageGuidelinesUrl = "";
        if (guidanceUrl !== null && guidanceUrl.length > 0) {
          siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl[0].value;
        }

        // Get the GuestUsageGuidelinesUrl
        const guestGuidanceUrl: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
          return directorySetting.name === 'GuestUsageGuidelinesUrl';
        });

        siteClassificationsSettings.GuestUsageGuidelinesUrl = "";
        if (guestGuidanceUrl !== null && guestGuidanceUrl.length > 0) {
          siteClassificationsSettings.GuestUsageGuidelinesUrl = guestGuidanceUrl[0].value;
        }

        // Get the DefaultClassification
        const defaultClassification: DirectorySettingValue[] = unifiedGroupSetting[0].values.filter((directorySetting: DirectorySettingValue): boolean => {
          return directorySetting.name === 'DefaultClassification';
        });

        siteClassificationsSettings.DefaultClassification = "";
        if (defaultClassification !== null && defaultClassification.length > 0) {
          siteClassificationsSettings.DefaultClassification = defaultClassification[0].value;
        }

        logger.log(JSON.parse(JSON.stringify(siteClassificationsSettings)));

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadSiteClassificationGetCommand();