import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import { DirectorySetting } from './DirectorySetting';
import { DirectorySettingTemplatesRsp } from './DirectorySettingTemplatesRsp';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadSiteClassificationDisableCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_DISABLE}`;
  }

  public get description(): string {
    return 'Disables site classification';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const disableSiteClassification: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/beta/settings`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .get<DirectorySettingTemplatesRsp>(requestOptions)
        .then((res: DirectorySettingTemplatesRsp): Promise<void> => {
          if (res.value.length === 0) {
            return Promise.reject('Site classification is not enabled.');
          }

          const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
            return directorySetting.displayName === 'Group.Unified';
          });

          if (!unifiedGroupSetting || unifiedGroupSetting.length === 0) {
            return Promise.reject('Missing DirectorySettingTemplate for "Group.Unified"');
          }

          if (!unifiedGroupSetting[0] ||
            !unifiedGroupSetting[0].id || unifiedGroupSetting[0].id.length === 0) {
            return Promise.reject('Missing UnifiedGroupSettting id');
          }

          const requestOptions: any = {
            url: `${this.resource}/beta/settings/` + unifiedGroupSetting[0].id,
            headers: {
              accept: 'application/json;odata.metadata=none',
              'content-type': 'application/json'
            },
            json: true,
          };

          return request.delete(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      disableSiteClassification();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disable site classification?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          disableSiteClassification();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming disabling site classification'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadSiteClassificationDisableCommand();