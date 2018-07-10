import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer,PlaceholderContent, PlaceholderName} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'SpfxTrinSuiteBarApplicationCustomizerStrings';
const LOG_SOURCE: string = 'SpfxTrinSuiteBarApplicationCustomizer';
import pnp, { PermissionKind } from "sp-pnp-js";
import styles from  "./AppCustomizer.module.scss";
import { SPComponentLoader } from '@microsoft/sp-loader';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxTrinSuiteBarApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxTrinSuiteBarApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxTrinSuiteBarApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    Dialog.alert(`Hello 123 from ${strings.Title}:\n\n${message}`);

    //Hide SuiteBat and Navigation Bar 
    this.HideSuiteBar();

    this.AddHeaderFooter();

    return Promise.resolve();


  }

  private HideSuiteBar() {
    var suiteBar = document.getElementById("SuiteNavPlaceHolder");
    suiteBar.setAttribute("style", "display: none");

    // var HubNavbigation = document.getElementsByClassName("ms-HubNav root-109")[0];
    // HubNavbigation.setAttribute("style", "display: none");

    var compositeHeader = document.getElementsByClassName("banner_6a0b822b")[0];
    compositeHeader.setAttribute("style", "display: none");

  }

   private AddHeaderFooter() {

    const HEADER_TEXT: string = "This is the top zone";
    const FOOTER_TEXT: string = "This is the bottom zone";

    //Top and Bottom Placeholder..
    let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    let bottomPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);

    // /*Top Place Holder */    
    if (topPlaceholder) {
      topPlaceholder.domElement.innerHTML =`<div class="${styles.header}">Custom Header Text</div>`;
    }

    /*Footer Place Holder */    
    if (bottomPlaceholder) {
      bottomPlaceholder.domElement.innerHTML =`<div class="${styles.footer}">Custom Footer Text</div>`;
    }

   }


}
