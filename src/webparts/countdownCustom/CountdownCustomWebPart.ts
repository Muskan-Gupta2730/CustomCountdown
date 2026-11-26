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
      domElement: this.domElement,
      onVisibilityChange: (visible: boolean) => this._setWebPartVisibility(visible),
    });
    ReactDom.render(element, this.domElement);
  }

  private _setWebPartVisibility(visible: boolean): void {
    if (this.displayMode === DisplayMode.Edit) {
      this._restoreVisibility();
      return;
    }

    if (!visible) {
      this.domElement.style.display = 'none';

      // 1. Find SharePoint web part control wrapper
      const controlWrapper = this.domElement.closest(
        '[data-sp-feature-tag="CanvasControl"], .CanvasControl, [data-automation-id="CanvasControl"], [data-sp-a11y-id], .ControlZone'
      ) as HTMLElement | null;

      if (controlWrapper) {
        controlWrapper.style.display = 'none';
      }

      // 2. Find SharePoint section wrapper
      const sectionWrapper = this.domElement.closest(
        '[data-sp-feature-tag="CanvasSection"], .CanvasSection, .CanvasZone, [data-automation-id="CanvasSection"]'
      ) as HTMLElement | null;

      if (sectionWrapper) {
        // Check if any other visible web part exists in this section
        const otherControls = Array.from(
          sectionWrapper.querySelectorAll('[data-sp-feature-tag="CanvasControl"], .CanvasControl, [data-automation-id="CanvasControl"]')
        ).filter(c => c !== controlWrapper && !c.contains(this.domElement) && (c as HTMLElement).style.display !== 'none');

        // If this is the only web part in the section, hide the whole section to eliminate empty space
        if (otherControls.length === 0) {
          sectionWrapper.style.display = 'none';
        }
      }
    } else {
      this._restoreVisibility();
    }
  }

  private _restoreVisibility(): void {
    this.domElement.style.display = '';

    const controlWrapper = this.domElement.closest(
      '[data-sp-feature-tag="CanvasControl"], .CanvasControl, [data-automation-id="CanvasControl"], [data-sp-a11y-id], .ControlZone'
    ) as HTMLElement | null;
    if (controlWrapper) {
      controlWrapper.style.display = '';
    }

    const sectionWrapper = this.domElement.closest(
      '[data-sp-feature-tag="CanvasSection"], .CanvasSection, .CanvasZone, [data-automation-id="CanvasSection"]'
    ) as HTMLElement | null;
    if (sectionWrapper) {
      sectionWrapper.style.display = '';
    }
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
    this._restoreVisibility();
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}