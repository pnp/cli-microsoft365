import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class SpeContainerGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_GET;
  }

  public get description(): string {
    return 'Gets a container of a specific container type';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' }
    );
  }

  #initTypes(): void {
    this.types.string.push('id');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting a container with id '${args.options.id}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/storage/fileStorage/containers/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerGetCommand();