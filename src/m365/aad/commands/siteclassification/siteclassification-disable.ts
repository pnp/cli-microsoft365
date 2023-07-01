import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const disableSiteClassification = async (): Promise<void> => {
      try {
        let requestOptions: CliRequestOptions = {
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

        if (!unifiedGroupSetting || unifiedGroupSetting.length === 0) {
          throw 'Missing DirectorySettingTemplate for "Group.Unified"';
        }

        if (!unifiedGroupSetting[0] ||
          !unifiedGroupSetting[0].id || unifiedGroupSetting[0].id.length === 0) {
          throw 'Missing UnifiedGroupSettting id';
        }

        requestOptions = {
          url: `${this.resource}/v1.0/groupSettings/` + unifiedGroupSetting[0].id,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await disableSiteClassification();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disable site classification?`
      });

      if (result.continue) {
        await disableSiteClassification();
      }
    }
  }
}

module.exports = new AadSiteClassificationDisableCommand();