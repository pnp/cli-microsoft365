import * as assert from 'assert';
import AadCommand from './AadCommand';

class MockCommand extends AadCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(): void {
  }

  public commandHelp(): void {
  }
}

describe('AadCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://graph.windows.net');
  });
});