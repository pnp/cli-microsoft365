export interface CommandInstance {
  commandWrapper: {
    command: string;
  };
  log: (message: any) => void;
  prompt: (object: any, callback: (result: any) => void) => void;
}