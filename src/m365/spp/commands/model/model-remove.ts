import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  id?: string;
  title?: string;
  force?: boolean;
}

class SppModelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.MODEL_REMOVE;
  }

  public get description(): string {
    return 'Deletes a document understanding model';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-f --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'id', 'title');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.force) {
        const confirmationResult = await cli.promptForConfirmation({ message: `Are you sure you want to remove the model '${args.options.id || args.options.title}'?` });

        if (!confirmationResult) {
          return;
        }
      }

      if (this.verbose) {
        await logger.log(`Removing model from ${args.options.siteUrl}...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      await spp.assertSiteIsContentCenter(siteUrl);

      const requestOptions: CliRequestOptions = {
        url: this.getCorrectRequestUrl(siteUrl, args),
        headers: {
          accept: 'application/json;odata=nometadata',
          'if-match': '*'
        }
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getCorrectRequestUrl(siteUrl: string, args: CommandArgs): string {
    if (args.options.id) {
      return `${siteUrl}/_api/machinelearning/models/getbyuniqueid('${args.options.id}')`;
    }

    return `${siteUrl}/_api/machinelearning/models/getbytitle('${formatting.encodeQueryParameter(args.options.title!)}')`;
  }
}

export default new SppModelRemoveCommand();