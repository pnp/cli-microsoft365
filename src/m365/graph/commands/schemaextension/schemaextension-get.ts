import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class GraphSchemaExtensionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_GET;
  }

  public get description(): string {
    return 'Gets the properties of the specified schema extension definition';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Gets the properties of the specified schema extension definition with id '${args.options.id}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}
module.exports = new GraphSchemaExtensionGetCommand();