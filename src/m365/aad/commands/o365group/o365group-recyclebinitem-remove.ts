import { Group } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  mailNickname?: string;
  confirm: boolean;
}

class AadO365GroupRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes a Microsoft 365 Group from the recycle bin in the current tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    telemetryProps.mailNickname = typeof args.options.mailNickname !== 'undefined';
    telemetryProps.confirm = !!args.options.confirm;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeGroup: () => void = async (): Promise<void> => {
      try {
        const groupId = await this.getGroupId(args.options);

        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/directory/deletedItems/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
        cb();
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      }
    };

    if (args.options.confirm) {
      removeGroup();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group '${args.options.id || args.options.displayName || args.options.mailNickname}'?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeGroup();
        }
      });
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    const { id, displayName, mailNickname } = options;

    if (id) {
      return id;
    }

    let filterValue: string = '';
    if (displayName) {
      filterValue = `displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
    }

    if (mailNickname) {
      filterValue = `mailNickname eq '${formatting.encodeQueryParameter(mailNickname)}'`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=${filterValue}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Group[] }>(requestOptions);
    const groups = response.value;

    if (groups.length === 0) {
      throw Error(`The specified group '${displayName || mailNickname}' does not exist.`);
    }

    if (groups.length > 1) {
      throw Error(`Multiple groups with name '${displayName || mailNickname}' found: ${groups.map(x => x.id).join(',')}.`);
    }

    return groups[0].id!;
  }

  public optionSets(): string[][] | undefined {
    return [['id', 'displayName', 'mailNickname']];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-d, --displayName [displayName]'
      },
      {
        option: '-m, --mailNickname [mailNickname]'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): string | boolean {
    if (args.options.id && !validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupRecycleBinItemRemoveCommand();