import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  allowUserEditMessages?: boolean;
  allowUserDeleteMessages?: boolean;
  allowOwnerDeleteMessages?: boolean;
  allowTeamMentions?: boolean;
  allowChannelMentions?: boolean;
}

class TeamsMessagingSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowUserEditMessages',
    'allowUserDeleteMessages',
    'allowOwnerDeleteMessages',
    'allowTeamMentions',
    'allowChannelMentions'
  ];

  public get name(): string {
    return commands.MESSAGINGSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates messaging settings of a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      TeamsMessagingSettingsSetCommand.booleanProps.forEach((p: string) => {
        this.telemetryProperties[p] = (args.options as any)[p];
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--allowUserEditMessages [allowUserEditMessages]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowUserDeleteMessages [allowUserDeleteMessages]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowOwnerDeleteMessages [allowOwnerDeleteMessages]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowTeamMentions [allowTeamMentions]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowChannelMentions [allowChannelMentions]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowUserEditMessages', 'allowUserDeleteMessages', 'allowOwnerDeleteMessages', 'allowTeamMentions', 'allowChannelMentions');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        let hasDuplicate: boolean = false;
        let property: string = '';
        TeamsMessagingSettingsSetCommand.booleanProps.forEach((prop: string) => {
          if ((args.options as any)[prop] instanceof Array) {
            property = prop;
            hasDuplicate = true;
          }
        });
        if (hasDuplicate) {
          return `Duplicate option ${property} specified. Specify only one`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      messagingSettings: {}
    };
    TeamsMessagingSettingsSetCommand.booleanProps.forEach((p: string) => {
      if (typeof (args.options as any)[p] !== 'undefined') {
        data.messagingSettings[p] = (args.options as any)[p];
      }
    });

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: data,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMessagingSettingsSetCommand();