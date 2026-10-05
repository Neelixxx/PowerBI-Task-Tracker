import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import RcaNews from './components/RcaNews';
import { IRcaNewsProps } from './components/IRcaNewsProps';

export interface IRcaNewsWebPartProps {
  title: string;
  intro: string;
  maxItems: number;
  categoryFilter: string;
}

export default class RcaNewsWebPart extends BaseClientSideWebPart<IRcaNewsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IRcaNewsProps> = React.createElement(RcaNews, {
      title: this.properties.title || 'RCA News & Updates',
      intro: this.properties.intro || 'The latest updates, announcements, and insights from Root Cause Analysis.',
      maxItems: this.properties.maxItems || 7,
      categoryFilter: this.properties.categoryFilter || '',
      context: this.context
    });
    ReactDom.render(element, this.domElement);
  }
  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Configure RCA News.' }, groups: [{
      groupName: 'Content',
      groupFields: [
        PropertyPaneTextField('title', { label: 'Section title' }),
        PropertyPaneTextField('intro', { label: 'Intro text', multiline: true }),
        PropertyPaneSlider('maxItems', { label: 'Number of posts', min: 4, max: 12, step: 1 }),
        PropertyPaneTextField('categoryFilter', { label: 'Optional category filter' })
      ]
    }]}]};
  }
}
