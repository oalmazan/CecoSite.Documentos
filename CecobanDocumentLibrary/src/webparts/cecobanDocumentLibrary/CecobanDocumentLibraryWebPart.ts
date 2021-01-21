import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CecobanDocumentLibraryWebPartStrings';
import { IFabricDetailsListProps } from './components/IFabricDetailsListProps';
import FabricDetailsList from './components/FabricDetailsList';

export interface ICecobanDocumentLibraryWebPartProps {
  limitDocuments: number;
  titleDocuments: string;
  documentsFilter: string;
  startFilesPath: string;
  description: string;
  filterUserProperty: string;
}

export default class CecobanDocumentLibraryWebPart extends BaseClientSideWebPart<ICecobanDocumentLibraryWebPartProps> {

  public render(): void {

    const element2: React.ReactElement<IFabricDetailsListProps > = React.createElement(
      FabricDetailsList,
      {
        spcontect: this.context,
        StartFilesPath: this.properties.startFilesPath,
        TitleDocuments: this.properties.titleDocuments,
        LimitDocuments: this.properties.limitDocuments,
        DocumentsFilter: this.properties.documentsFilter,
        UserName: encodeURIComponent('i:0#.f|membership|' + this.context.pageContext.user.loginName),
        ServiceScope: this.context.serviceScope,
        FilterProperty: this.properties.filterUserProperty
      }
    );

    ReactDom.render(element2, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('filterUserProperty', {
                  label: strings.FilterUserPropertyFieldLabel
                }),
                PropertyPaneTextField('limitDocuments', {
                  label: strings.LimitDocumentsPropertiesFieldLabel
                }),
                PropertyPaneTextField('documentsFilter', {
                  label: strings.DocumentsFilterPropertiesFieldLabel
                }),
                PropertyPaneTextField('titleDocuments', {
                  label: strings.TitleDocumentsPropertyFieldLabel
                }),
                PropertyPaneTextField('startFilesPath', {
                  label: strings.StartFilesPathPropertyFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
