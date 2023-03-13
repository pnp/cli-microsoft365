import { Logger } from '../../../cli/Logger';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

class MockCommandWithOptionSets extends AnonymousCommand {
  public get name(): string {
    return commands.TEST;
  }
  public get description(): string {
    return 'Mock command with option sets';
  }
  constructor() {
    super();

    this.options.push(
      {
        option: '--opt1 [name]'
      },
      {
        option: '--opt2 [name]'
      },
      {
        option: '--opt3 [name]'
      },
      {
        option: '--opt4 [name]'
      },
      {
        option: '--opt5 [name]'
      },
      {
        option: '--opt6 [name]'
      }
    );
    this.optionSets.push(
      {
        options: ['opt5', 'opt6'],
        runsWhen: (args) => typeof args.options.opt4 !== 'undefined'  // validate when opt4 is passed
      },
      {
        options: ['opt3', 'opt4'],
        runsWhen: (args) => typeof args.options.opt2 !== 'undefined' // validate when opt2 is passed
      },
      { options: ['opt1', 'opt2'] }
    );
  }
  public async commandAction(logger: Logger): Promise<void> {
    logger.log('');
  }
}
module.exports = new MockCommandWithOptionSets();