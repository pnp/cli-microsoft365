export interface Message {
  body: {
    content: string;
  };
  id: string;
  summary: string | null;
}