interface Omelette {
  init: () => never;
  on: (event: string, eventHandler: (fragment: string, data: EventData) => void) => void;
  setupShellInitFile: (initFile?: string) => void;
}

interface EventData {
  before: any;
  fragment: number;
  line: string;
  reply: (words: any) => never;
}