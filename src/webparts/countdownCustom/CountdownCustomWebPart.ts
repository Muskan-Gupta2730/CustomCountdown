import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

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
    p.listTitle        = p.listTitle        || 'CountdownEvents';
    p.label            = p.label            || 'Application closes in';
    p.buttonText       = p.buttonText       || 'Apply now';
    p.description      = p.description      || '';
    p.expiredMessage   = p.expiredMessage   || 'Applications are now closed';
    p.overlayOpacity   = p.overlayOpacity   ?? 0;
    p.backgroundImageUrl = p.backgroundImageUrl || '';
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}