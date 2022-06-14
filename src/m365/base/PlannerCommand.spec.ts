import * as assert from 'assert';
import PlannerCommand from './PlannerCommand';

class MockCommand extends PlannerCommand {
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

describe('PlannerCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://tasks.office.com');
  });
});