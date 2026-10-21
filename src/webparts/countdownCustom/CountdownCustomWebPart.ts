import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import CountdownCustom from './components/CountdownCustom';
import {
  ICountdownCustomWebPartProps,
  ICountdownCustomProps,
} from './components/ICountdownCustomProps';

export default class CountdownCustomWebPart extends BaseClientSideWebPart<ICountdownCustomWebPartProps> {
  private _listOptions: IPropertyPaneDropdownOption[] = [
    { key: 'CountdownEvents', text: 'CountdownEvents' },
  ];
  private _listsFetched: boolean = false;
  private _listsLoading: boolean = false;

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
    p.listTitle        = p.listTitle        || 'CountdownEvents';
    p.label            = p.label            || 'Application closes in';
    p.buttonText       = p.buttonText       || 'Apply now';
    p.description      = p.description      || '';
    p.expiredMessage   = p.expiredMessage   || 'Applications are now closed';
    p.overlayOpacity   = p.overlayOpacity   ?? 0;
    p.backgroundImageUrl = p.backgroundImageUrl || '';
    return Promise.resolve();
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.properties.listTitle && !this._listOptions.some((o) => o.key === this.properties.listTitle)) {
      this._listOptions.unshift({ key: this.properties.listTitle, text: this.properties.listTitle });
    }
    this._loadLists();
  }

  private _loadLists(): void {
    if (this._listsFetched || this._listsLoading) return;
    this._listsLoading = true;

    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Title&$orderby=Title`;
    this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
      })
      .then((data: { value: Array<{ Title: string }> }) => {
        const fetched: IPropertyPaneDropdownOption[] = (data.value || []).map((l) => ({
          key: l.Title,
          text: l.Title,
        }));
        if (this.properties.listTitle && !fetched.some((o) => o.key === this.properties.listTitle)) {
          fetched.unshift({ key: this.properties.listTitle, text: this.properties.listTitle });
        }
        if (fetched.length > 0) {
          this._listOptions = fetched;
        }
        this._listsFetched = true;
        this._listsLoading = false;
        this.context.propertyPane.refresh();
      })
      .catch((err) => {
        console.warn('Could not load SharePoint lists for property pane dropdown:', err);
        this._listsFetched = true;
        this._listsLoading = false;
      });
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
              groupName: 'List configuration',
              groupFields: [
                PropertyPaneDropdown('listTitle', {
                  label: 'Select list',
                  options: this._listOptions,
                  selectedKey: this.properties.listTitle,
                }),
                PropertyPaneTextField('description', {
                  label: 'Description text (fallback if not in list)',
                  multiline: true,
                  rows: 3,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}