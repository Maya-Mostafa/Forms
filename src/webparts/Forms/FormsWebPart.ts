import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FormsWebPartStrings';
import Forms from './components/Forms';
import { IFormsProps } from './components/IFormsProps';

export interface IFormsWebPartProps {
  wpTitle: string;
  listUrl: string;
  listName: string;
  pageSize: number;
  showHelp: boolean;
  helpLink: string;
  helpTitle: string;
  showHelpMsg: boolean;
  helpMsgTxt: string;
  helpMsgLink: string;
  helpMsgLinkTxt: string;
}

export default class FormsWebPart extends BaseClientSideWebPart<IFormsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFormsProps> = React.createElement(
      Forms,
      {
        context: this.context,
        wpTitle: this.properties.wpTitle,
        listUrl: this.properties.listUrl,
        listName: this.properties.listName,
        pageSize: this.properties.pageSize,
        helpLink: this.properties.helpLink,
        helpTitle: this.properties.helpTitle,
        showHelp: this.properties.showHelp,
        showHelpMsg: this.properties.showHelpMsg,
        helpMsgTxt: this.properties.helpMsgTxt,
        helpMsgLink: this.properties.helpMsgLink,
        helpMsgLinkTxt: this.properties.helpMsgLinkTxt
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('wpTitle', {
                  label: 'Webpart Title',
                  value: this.properties.wpTitle
                }),
                PropertyPaneTextField('listUrl', {
                  label: 'List URL',
                  value: this.properties.listUrl
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name',
                  value: this.properties.listName
                }),
                PropertyPaneTextField('pageSize', {
                  label: 'Number of Items',
                  value: this.properties.pageSize.toString()
                })
              ]
            },
            {
              groupName: "Help Icon",
              groupFields: [
                PropertyPaneCheckbox('showHelp', {
                  text: "Show help icon",
                  checked: this.properties.showHelp
                }),
                PropertyPaneTextField('helpTitle', {
                  label: 'Link Text',
                  value: this.properties.helpTitle
                }),
                PropertyPaneTextField('helpLink', {
                  label: 'Link',
                  value: this.properties.helpLink
                }),
              ]
            },
            {
              groupName: "Help Message",
              groupFields: [
                PropertyPaneCheckbox('showHelpMsg', {
                  text: "Show help message",
                  checked: this.properties.showHelpMsg
                }),
                PropertyPaneTextField('helpMsgTxt', {
                  label: 'Message Text',
                  value: this.properties.helpMsgTxt
                }),
                PropertyPaneTextField('helpMsgLinkTxt', {
                  label: 'Link Text',
                  value: this.properties.helpMsgLinkTxt
                }),
                PropertyPaneTextField('helpMsgLink', {
                  label: 'Link',
                  value: this.properties.helpMsgLink
                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
