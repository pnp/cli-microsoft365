import Command from '../../Command';

export interface GraphResponseError {
  error: {
    code: string;
    message: string;
    innerError: {
      "request-id": string;
      date: string;
    }
  }
}

export default abstract class GraphCommand extends Command {
  protected get resource(): string {
    return 'https://graph.microsoft.com';
  }
}