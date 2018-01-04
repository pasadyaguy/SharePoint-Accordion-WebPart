import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import AccordionTemplate from "./AccordionTemplate";
import styles from "./AccordionWebPart.module.scss";
import * as strings from "AccordionWebPartStrings";
import * as jQuery from "jquery";
import "jqueryui";
import  { SPComponentLoader } from "@microsoft/sp-loader";


export interface IAccordionWebPartProps {
  description: string;

}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {

  public constructor() {
    super();

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css");
  }

  public render(): void {
    this.domElement.innerHTML = AccordionTemplate.templateHtml;

    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: false,
      icons: {
        header: "ui-icon-circle-arrow-e",
        activeHeader: "ui-icon-circle-arrow-s"
      }
    };

    jQuery(".accordion", this.domElement).accordion(accordionOptions);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
