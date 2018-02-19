// import * as assert from 'assert';
// import { GraphItemsListCommand } from './GraphItemsListCommand';

// class Item {}

// class MockCommand extends GraphItemsListCommand<Item> {
//   public get name(): string {
//     return 'mock';
//   }

//   public get description(): string {
//     return 'Mock command';
//   }

//   public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
//     cb();
//   }

//   public commandHelp(args: any, log: (message: string) => void): void {
//   }
// }

// describe('GraphItemsListCommand', () => {
//   it('initializes default collection of items', () => {
//     const command = new MockCommand();
//     assert.deepEqual((command as any).items, []);
//   });
// });