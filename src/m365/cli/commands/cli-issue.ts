import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(['bug', 'command', 'sample']).alias('t')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliIssueCommand extends AnonymousCommand {
  public get name(): string {
    return commands.ISSUE;
  }

  public get description(): string {
    return 'Returns, or opens a URL that takes the user to the right place in the CLI GitHub repo to create a new issue reporting bug, feedback, ideas, etc.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let issueLink: string = '';

    switch (encodeURIComponent(args.options.type)) {
      case 'bug':
        issueLink = 'https://aka.ms/cli-m365/bug';
        break;
      case 'command':
        issueLink = 'https://aka.ms/cli-m365/new-command';
        break;
      case 'sample':
        issueLink = 'https://aka.ms/cli-m365/new-sample-script';
        break;
    }

    await browserUtil.open(issueLink);
    await logger.log(issueLink);
  }
}

export default new CliIssueCommand();
