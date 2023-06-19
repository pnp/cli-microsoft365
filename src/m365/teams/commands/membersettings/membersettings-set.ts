import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  allowAddRemoveApps?: boolean;
  allowCreateUpdateChannels?: boolean;
  allowCreateUpdateRemoveConnectors?: boolean;
  allowCreateUpdateRemoveTabs?: boolean;
  allowDeleteChannels?: boolean;
  teamId: string;
}

class TeamsMemberSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowAddRemoveApps',
    'allowCreateUpdateChannels',
    'allowCreateUpdateRemoveConnectors',
    'allowCreateUpdateRemoveTabs',
    'allowDeleteChannels'
  ];

  public get name(): string {
    return commands.MEMBERSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates member settings of a Microsoft Teams team';
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
      TeamsMemberSettingsSetCommand.booleanProps.forEach(p => {
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
        option: '--allowAddRemoveApps [allowAddRemoveApps]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowCreateUpdateChannels [allowCreateUpdateChannels]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowCreateUpdateRemoveConnectors [allowCreateUpdateRemoveConnectors]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowCreateUpdateRemoveTabs [allowCreateUpdateRemoveTabs]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowDeleteChannels [allowDeleteChannels]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowAddRemoveApps', 'allowCreateUpdateChannels', 'allowCreateUpdateRemoveConnectors', 'allowCreateUpdateRemoveTabs', 'allowDeleteChannels');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      memberSettings: {}
    };
    TeamsMemberSettingsSetCommand.booleanProps.forEach(p => {
      if (typeof (args.options as any)[p] !== 'undefined') {
        data.memberSettings[p] = (args.options as any)[p];
      }
    });

    const requestOptions: CliRequestOptions = {
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

module.exports = new TeamsMemberSettingsSetCommand();