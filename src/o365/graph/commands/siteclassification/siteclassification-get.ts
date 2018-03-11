import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';

import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { DirectorySettingTemplatesRsp } from './DirectorySettingTemplatesRsp';
import { DirectorySetting } from './DirectorySetting';
import { DirectorySettingValue } from './DirectorySettingValue';
import { SiteClassificationSettings} from './SiteClassificationSettings'
import { CommandError } from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
}

class GraphO365SiteClassificationGetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_GET}`;
  }

  public get description(): string {
    return 'Get site classification configuration';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: DirectorySettingTemplatesRsp): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        if(res.value.length == 0) { 
          cmd.log(new CommandError('SiteClassification is not enabled.'));
        }
        else{
          var unifiedGroupSetting = res.value.filter(function (directorySetting : DirectorySetting) {
            return directorySetting.displayName === 'Group.Unified';
          });
          
          if(unifiedGroupSetting == null || unifiedGroupSetting.length == 0)
          {
            cmd.log(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
          }
          else {
            var siteClassificationsSettings  = new SiteClassificationSettings();

            // Get the classification list
            var classificationList = unifiedGroupSetting[0].values.filter(function (directorySetting : DirectorySettingValue) {
              return directorySetting.name === 'ClassificationList';
            });

            siteClassificationsSettings.Classifications = [];
            if (classificationList != null && classificationList.length > 0)
            {
                siteClassificationsSettings.Classifications = classificationList[0].value.split(',');
            }

            // Get the UsageGuidancelinesUrl
            var guidanceUrl = unifiedGroupSetting[0].values.filter(function (directorySetting : DirectorySettingValue) {
              return directorySetting.name === 'UsageGuidelinesUrl';
            });

            siteClassificationsSettings.UsageGuidelinesUrl = "";
            if (guidanceUrl != null && guidanceUrl.length > 0)
            {
                siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl[0].value;
            }

            // Get the DefaultClassification
            var defaultClassification = unifiedGroupSetting[0].values.filter(function (directorySetting : DirectorySettingValue) {
              return directorySetting.name === 'DefaultClassification';
            });

            siteClassificationsSettings.DefaultClassification = "";
            if (defaultClassification != null && defaultClassification.length > 0)
            {
                siteClassificationsSettings.DefaultClassification = defaultClassification[0].value;
            }

            cmd.log(siteClassificationsSettings);
          }

        }

        cb();
      }, (err: any) => 
      this.handleRejectedODataJsonPromise(err, cmd, cb));
  }


  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To get information about a Office 365 Tenant SiteClassification, you have to first connect to
    the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

  Examples:
  
    Get information about the Office 365 Tenant SiteClassification}
      ${chalk.grey(config.delimiter)} ${this.name}
    `);
  }
}

module.exports = new GraphO365SiteClassificationGetCommand();