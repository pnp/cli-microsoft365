import * as assert from 'assert';
import { Logger } from '../../cli';
import AadCommand from './AadCommand';

class MockCommand extends AadCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(logger: Logger, args: {}, cb: () => void): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('AadCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://graph.windows.net');
  });
});