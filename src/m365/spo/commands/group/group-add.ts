import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  description?: string;
  allowMembersEditMembership?: boolean;
  onlyAllowMembersViewMembership?: boolean;
  allowRequestToJoinLeave?: boolean;
  autoAcceptRequestToJoinLeave?: boolean;
  requestToJoinLeaveEmailSetting?: string;
}

class SpoGroupAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_ADD;
  }

  public get description(): string {
    return 'Creates group in the specified site';
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
        description: typeof args.options.description !== 'undefined',
        allowMembersEditMembership: typeof args.options.allowMembersEditMembership !== 'undefined',
        onlyAllowMembersViewMembership: typeof args.options.onlyAllowMembersViewMembership !== 'undefined',
        allowRequestToJoinLeave: typeof args.options.allowRequestToJoinLeave !== 'undefined',
        autoAcceptRequestToJoinLeave: typeof args.options.autoAcceptRequestToJoinLeave !== 'undefined',
        requestToJoinLeaveEmailSetting: typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
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
      url: `${args.options.webUrl}/_api/web/sitegroups`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        Title: args.options.name,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    request
      .post(requestOptions)
      .then((response: any): void => {
        logger.log(response);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoGroupAddCommand();