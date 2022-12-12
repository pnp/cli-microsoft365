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
  allowGiphy?: boolean;
  giphyContentRating: string;
  allowStickersAndMemes?: boolean;
  allowCustomMemes?: boolean;
}

class TeamsFunSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowGiphy',
    'allowStickersAndMemes',
    'allowCustomMemes'
  ];

  public get name(): string {
    return commands.FUNSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates fun settings of a Microsoft Teams team';
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
      Object.assign(this.telemetryProperties, {
        giphyContentRating: args.options.giphyContentRating
      });
      TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
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
        option: '--allowGiphy [allowGiphy]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--giphyContentRating [giphyContentRating]'
      },
      {
        option: '--allowStickersAndMemes [allowStickersAndMemes]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowCustomMemes [allowCustomMemes]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('allowGiphy', 'allowStickersAndMemes', 'allowCustomMemes');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.giphyContentRating) {
          const giphyContentRating = args.options.giphyContentRating.toLowerCase();
          if (giphyContentRating !== 'strict' && giphyContentRating !== 'moderate') {
            return `giphyContentRating value ${args.options.giphyContentRating} is not valid.  Please specify Strict or Moderate.`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const data: any = {
        funSettings: {}
      };
      TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
        if (typeof (args.options as any)[p] !== 'undefined') {
          data.funSettings[p] = (args.options as any)[p];
        }
      });

      if (args.options.giphyContentRating) {
        data.funSettings.giphyContentRating = args.options.giphyContentRating;
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsFunSettingsSetCommand();