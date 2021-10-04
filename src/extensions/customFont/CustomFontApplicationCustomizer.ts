import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CustomFontApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CustomFontApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomFontApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  siteurl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomFontApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomFontApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    this.properties.siteurl = this.context.pageContext.site.absoluteUrl;
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const cssUrl: string = this.properties.siteurl + "/Shared%20Documents/Styling/IntranetStyling.css";
    if (cssUrl) {
      // inject the style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }

    return Promise.resolve();
  }
}
