import { AxiosRequestConfig } from 'axios';
import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  newName?: string;
  description?: string;
  allowMembersEditMembership?: boolean;
  onlyAllowMembersViewMembership?: boolean;
  allowRequestToJoinLeave?: boolean;
  autoAcceptRequestToJoinLeave?: boolean;
  requestToJoinLeaveEmailSetting?: boolean;
  ownerEmail?: string;
  ownerUserName?: string;
}

class SpoGroupSetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_SET;
  }

  public get description(): string {
    return 'Updates a group in the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        allowMembersEditMembership: typeof args.options.allowMembersEditMembership !== 'undefined',
        onlyAllowMembersViewMembership: typeof args.options.onlyAllowMembersViewMembership !== 'undefined',
        allowRequestToJoinLeave: typeof args.options.allowRequestToJoinLeave !== 'undefined',
        autoAcceptRequestToJoinLeave: typeof args.options.autoAcceptRequestToJoinLeave !== 'undefined',
        requestToJoinLeaveEmailSetting: typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined',
        ownerEmail: typeof args.options.ownerEmail !== 'undefined',
        ownerUserName: typeof args.options.ownerUserName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--newName [newName]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--allowMembersEditMembership [allowMembersEditMembership]'
      },
      {
        option: '--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]'
      },
      {
        option: '--allowRequestToJoinLeave [allowRequestToJoinLeave]'
      },
      {
        option: '--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]'
      },
      {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
      },
      {
        option: '--ownerEmail [ownerEmail]'
      },
      {
        option: '--ownerUserName [ownerUserName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id && isNaN(args.options.id)) {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.ownerEmail && args.options.ownerUserName) {
          return 'Specify either ownerEmail or ownerUserName but not both';
        }

        const booleanOptions = [
          args.options.allowMembersEditMembership, args.options.onlyAllowMembersViewMembership,
          args.options.allowRequestToJoinLeave, args.options.autoAcceptRequestToJoinLeave
        ];
        for (const option of booleanOptions) {
          if (typeof option !== 'undefined' && !validation.isValidBoolean(option as any)) {
            return `Value '${option}' is not a valid boolean`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: AxiosRequestConfig = {
      url: `${args.options.webUrl}/_api/web/sitegroups/${args.options.id ? `GetById(${args.options.id})` : `GetByName('${args.options.name}')`}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        Title: args.options.newName,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    request
      .patch(requestOptions)
      .then(() => this.setGroupOwner(args.options))
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private setGroupOwner(options: Options): Promise<void> {
    if (!options.ownerEmail && !options.ownerUserName) {
      return Promise.resolve();
    }

    return this
      .getOwnerId(options)
      .then((ownerId: number): Promise<void> => {
        const requestOptions: AxiosRequestConfig = {
          url: `${options.webUrl}/_api/web/sitegroups/${options.id ? `GetById(${options.id})` : `GetByName('${options.name}')`}/SetUserAsOwner(${ownerId})`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      });
  }

  private getOwnerId(options: Options): Promise<number> {
    const cmdOptions: AadUserGetCommandOptions = {
      userName: options.ownerUserName,
      email: options.ownerEmail,
      output: 'json',
      debug: options.debug,
      verbose: options.verbose
    };

    return Cli
      .executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...cmdOptions, _: [] } })
      .then((output: CommandOutput) => {
        const getUserOutput = JSON.parse(output.stdout);

        const requestOptions: AxiosRequestConfig = {
          url: `${options.webUrl}/_api/web/ensureUser('${getUserOutput.userPrincipalName}')?$select=Id`,
          headers: {
            accept: 'application/json',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        return request.post<{ Id: number }>(requestOptions);
      })
      .then((response: { Id: number }): number => response.Id);
  }
}

module.exports = new SpoGroupSetCommand();