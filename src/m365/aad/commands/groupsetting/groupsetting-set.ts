import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { GroupSetting } from './GroupSetting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class AadGroupSettingSetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_SET;
  }

  public get description(): string {
    return 'Updates the particular group setting';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving group setting with id '${args.options.id}'...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<GroupSetting>(requestOptions)
      .then((groupSetting: GroupSetting): Promise<void> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          data: {
            displayName: groupSetting.displayName,
            templateId: groupSetting.templateId,
            values: this.getGroupSettingValues(args.options, groupSetting)
          },
          responseType: 'json'
        };

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getGroupSettingValues(options: any, groupSetting: GroupSetting): { name: string; value: string }[] {
    const values: { name: string; value: string }[] = [];
    const excludeOptions: string[] = [
      'id',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        values.push({
          name: key,
          value: options[key]
        });
      }
    });

    groupSetting.values.forEach(v => {
      if (!values.find(e => e.name === v.name)) {
        values.push({
          name: v.name,
          value: v.value
        });
      }
    });

    return values;
  }
}

module.exports = new AadGroupSettingSetCommand();