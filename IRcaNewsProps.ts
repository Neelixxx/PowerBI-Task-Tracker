import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IRcaNewsProps {
  title: string;
  intro: string;
  maxItems: number;
  categoryFilter: string;
  context: WebPartContext;
}
