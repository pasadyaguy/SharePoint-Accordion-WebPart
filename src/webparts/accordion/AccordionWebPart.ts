import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import {
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";

import AccordionTemplate from "./AccordionTemplate";
import styles from "./AccordionWebPart.module.scss";
import * as strings from "AccordionWebPartStrings";
import * as jQuery from "jquery";
import "jqueryui";
import  { SPComponentLoader } from "@microsoft/sp-loader";
import MockHttpClient from "./MockHttpClient";
import { IODataList } from "@microsoft/sp-odata-types";

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface ISPListItems {
  value: ISPListItem[];
}

export interface ISPListItem {
  Title: string;
  Description: string;
}

export interface IAccordionWebPartProps {
  description: string;
  listName: string;
  sortOrder: string;
  lists: ISPLists[];
  dropdownOptions: IPropertyPaneDropdownOption[];
  sortOrderOptions: IPropertyPaneDropdownOption[];
}

export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {

  public constructor() {
    super();

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css");
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div id="spListItems" class="accordion">
      </div>`;

    this._renderSortOrder();
    this._renderListAsync();
    this._renderListItemAsync();
  }

  // x Sort Order Dropdown
  private _renderSortOrder(): void {
    this._getSortOrder()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.properties.sortOrderOptions = listOptions;
        this.context.propertyPane.refresh();
      });
  }

  private _getSortOrder(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void,
    reject: (error: any) => void) => {
      setTimeout((): void => {
        resolve([{
          key: "asc",
          text: "Ascending"
        },
        {
          key: "desc",
          text: "Descending"
        }]);
      }, 2000);
    });
  }

  // x SP Lists
  private _renderListAsync(): void {
    // local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    } else if (Environment.type === EnvironmentType.SharePoint ||
              Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
          console.log(response.value);
        });
    }
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }

  private _renderList(items: ISPList[]): void {
    var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    items.forEach((item: ISPList) => {
      options.push( { key: item.Title, text: item.Title });
    });

    this.properties.dropdownOptions = options;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
      + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // x SP List Items
  private _renderListItemAsync(): void {
    // local environment
    if (Environment.type === EnvironmentType.Local) {
      const container: Element = this.domElement.querySelector("#spListItems");
      container.innerHTML = `
      <h3>Section 1</h3>
      <div>
          <p>
          Mauris mauris ante, blandit et, ultrices a, suscipit eget, quam. Integer
          ut neque. Vivamus nisi metus, molestie vel, gravida in, condimentum sit
          amet, nunc. Nam a nibh. Donec suscipit eros. Nam mi. Proin viverra leo ut
          odio. Curabitur malesuada. Vestibulum a velit eu ante scelerisque vulputate.
          </p>
      </div>
      <h3>Section 2</h3>
      <div>
          <p>
          Sed non urna. Donec et ante. Phasellus eu ligula. Vestibulum sit amet
          purus. Vivamus hendrerit, dolor at aliquet laoreet, mauris turpis porttitor
          velit, faucibus interdum tellus libero ac justo. Vivamus non quam. In
          suscipit faucibus urna.
          </p>
      </div>
      <h3>Section 3</h3>
      <div>
          <p>
          Nam enim risus, molestie et, porta ac, aliquam ac, risus. Quisque lobortis.
          Phasellus pellentesque purus in massa. Aenean in pede. Phasellus ac libero
          ac tellus pellentesque semper. Sed ac felis. Sed commodo, magna quis
          lacinia ornare, quam ante aliquam nisi, eu iaculis leo purus venenatis dui.
          </p>
          <ul>
          <li>List item one</li>
          <li>List item two</li>
          <li>List item three</li>
          </ul>
      </div>
      <h3>Section 4</h3>
      <div>
          <p>
          Cras dictum. Pellentesque habitant morbi tristique senectus et netus
          et malesuada fames ac turpis egestas. Vestibulum ante ipsum primis in
          faucibus orci luctus et ultrices posuere cubilia Curae; Aenean lacinia
          mauris vel est.
          </p>
          <p>
          Suspendisse eu nisl. Nullam ut libero. Integer dignissim consequat lectus.
          Class aptent taciti sociosqu ad litora torquent per conubia nostra, per
          inceptos himenaeos.
          </p>
      </div>`;

      this.initAccordion();

    } else if (Environment.type === EnvironmentType.SharePoint ||
               Environment.type === EnvironmentType.ClassicSharePoint) {
      this._getListItemData()
          .then((response) => {
            this._renderListItem(response.value);
            console.log(response.value);
          });
    }
  }

  private _getListItemData(): Promise<ISPListItems> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
      + `/_api/web/lists/GetByTitle('${ this.properties.listName }')/items?$orderby=Title ${ this.properties.sortOrder }`,
          SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderListItem(items: ISPListItem[]): void {
    const container: Element = this.domElement.querySelector("#spListItems");
    let html: string = "";
    items.forEach((item: ISPListItem) => {
      html += `
      <h3>${ item.Title }</h3>
      <div>
          <p>
          ${ item.Description }
          </p>
      </div>`;
    });

    container.innerHTML = html;

    this.initAccordion();
  }

  private initAccordion(): void {
    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: true,
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
                PropertyPaneDropdown("listName", {
                  label: "Select a list",
                  options: this.properties.dropdownOptions,
                  disabled: false
                }),
                PropertyPaneDropdown("sortOrder", {
                  label: "Sort Order",
                  options: this.properties.sortOrderOptions,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
