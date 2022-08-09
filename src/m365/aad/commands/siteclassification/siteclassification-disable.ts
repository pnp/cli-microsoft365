import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { DirectorySetting } from './DirectorySetting';
import { DirectorySettingTemplatesRsp } from './DirectorySettingTemplatesRsp';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadSiteClassificationDisableCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_DISABLE;
  }

  public get description(): string {
    return 'Disables site classification';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const disableSiteClassification: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/groupSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
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
            url: `${this.resource}/v1.0/groupSettings/` + unifiedGroupSetting[0].id,
            headers: {
              accept: 'application/json;odata.metadata=none',
              'content-type': 'application/json'
            },
            responseType: 'json'
          };

          return request.delete(requestOptions);
        })
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      disableSiteClassification();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disable site classification?`
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
}

module.exports = new AadSiteClassificationDisableCommand();