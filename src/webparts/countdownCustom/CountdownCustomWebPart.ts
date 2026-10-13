import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  PropertyFieldFilePicker,
  IFilePickerResult,
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import CountdownCustom from './components/CountdownCustom';
import {
  ICountdownCustomWebPartProps,
  ICountdownCustomProps,
} from './components/ICountdownCustomProps';

export default class CountdownCustomWebPart extends BaseClientSideWebPart<ICountdownCustomWebPartProps> {

  public render(): void {
    const element = React.createElement<ICountdownCustomProps>(CountdownCustom, {
      ...this.properties,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      isEditMode: this.displayMode === DisplayMode.Edit,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    const p = this.properties;
    p.listTitle        = p.listTitle        ?? 'CountdownEvents';
    p.label            = p.label            ?? 'Applications close in';
    p.buttonText       = p.buttonText       ?? 'Apply now';
    p.description      = p.description      ?? '';
    p.expiredMessage   = p.expiredMessage   ?? 'Applications are now closed';
    p.overlayOpacity   = p.overlayOpacity   ?? 85;
    p.backgroundImageUrl = p.backgroundImageUrl ?? '';
    return Promise.resolve();
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
          header: { description: 'Countdown settings' },
          groups: [
            {
              groupName: 'Data source & text',
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: 'SharePoint list name',
                  description: 'Columns: Title, DateTime, Description, ApplyLink, isActive, PageLinks',
                }),
                PropertyPaneTextField('label', { label: 'Small heading (e.g. "Applications close in")' }),
                PropertyPaneTextField('description', {
                  label: 'Description text (fallback if not in list)',
                  multiline: true,
                  rows: 3,
                }),
                PropertyPaneTextField('buttonText', { label: 'Button text (e.g. "Apply now")' }),
                PropertyPaneTextField('expiredMessage', { label: 'Message after countdown ends' }),
              ],
            },
            {
              groupName: 'Background image',
              groupFields: [
                PropertyFieldFilePicker('backgroundImage', {
                  context: this.context as any,
                  filePickerResult: undefined as unknown as IFilePickerResult,
                  buttonLabel: 'Upload / choose image',
                  accepts: ['.png', '.jpg', '.jpeg', '.webp', '.svg'],
                  hideLocalUploadTab: false,
                  hideSiteFilesTab: false,
                  hideStockImages: true,
                  hideOneDriveTab: true,
                  hideWebSearchTab: true,
                  hideLinkUploadTab: false,
                  hideRecentTab: false,
                  hideOrganisationalAssetTab: true,
                  buttonIcon: 'FileImage',
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  onSave: (r: IFilePickerResult) => {
                    this.properties.backgroundImageUrl = r.fileAbsoluteUrl || '';
                    this.render();
                    this.context.propertyPane.refresh();
                  },
                  onChanged: (r: IFilePickerResult) => {
                    this.properties.backgroundImageUrl = r.fileAbsoluteUrl || '';
                    this.render();
                    this.context.propertyPane.refresh();
                  },
                  key: 'bgImageField',
                }),
                PropertyPaneTextField('backgroundImageUrl', {
                  label: 'Or paste image URL (clear to remove)',
                }),
                PropertyPaneSlider('overlayOpacity', {
                  label: 'Panel colour overlay over image (%)',
                  min: 0,
                  max: 100,
                  step: 5,
                  showValue: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}