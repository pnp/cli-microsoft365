import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
    this.#initTypes();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        description: typeof args.options.description !== 'undefined',
        allowMembersEditMembership: args.options.allowMembersEditMembership,
        onlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        allowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        autoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
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
        option: '--allowMembersEditMembership [allowMembersEditMembership]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowRequestToJoinLeave [allowRequestToJoinLeave]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowMembersEditMembership', 'onlyAllowMembersViewMembership', 'allowRequestToJoinLeave', 'autoAcceptRequestToJoinLeave');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
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

    try {
      const response = await request.post<any>(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoGroupAddCommand();