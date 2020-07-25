import * as assert from 'assert';
import AzmgmtCommand from './AzmgmtCommand';
import { CommandInstance } from '../../cli';

class MockCommand extends AzmgmtCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
  }

  public commandHelp(args: any, log: (message: string) => void): void {
  }
}

describe('AzmgmtCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://management.azure.com/');
  });
});