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
  id?: string;
  name?: string;
  description?: string;
  mailNickName?: string;
  classification?: string;
  visibility?: string;
}

class TeamsTeamSetCommand extends GraphCommand {
  private static props: string[] = [
    'description',
    'mailNickName',
    'classification',
    'visibility '
  ];

  public get name(): string {
    return commands.TEAM_SET;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      TeamsTeamSetCommand.props.forEach((p: string) => {
        this.telemetryProperties[p] = typeof (args.options as any)[p] !== 'undefined';
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--mailNickName [mailNickName]'
      },
      {
        option: '--classification [classification]'
      },
      {
        option: '--visibility [visibility]',
        autocomplete: ['Private', 'Public']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.visibility) {
          if (args.options.visibility.toLowerCase() !== 'private' && args.options.visibility.toLowerCase() !== 'public') {
            return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
          }
        }

        return true;
      }
    );
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.name) {
      requestBody.displayName = options.name;
    }
    if (options.description) {
      requestBody.description = options.description;
    }
    if (options.mailNickName) {
      requestBody.mailNickName = options.mailNickName;
    }
    if (options.classification) {
      requestBody.classification = options.classification;
    }
    if (options.visibility) {
      requestBody.visibility = options.visibility;
    }
    return requestBody;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(args.options.id as string)}`,
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

module.exports = new TeamsTeamSetCommand();