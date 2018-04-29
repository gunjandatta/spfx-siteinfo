import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteInfoWebPart.module.scss';
import * as strings from 'SiteInfoWebPartStrings';

// Import the gd-sprest libraries
import { Types } from "gd-sprest";
import { $REST, Fabric } from "gd-sprest-js";
import "gd-sprest-js/build/lib/css/fabric.components.min.css";

export interface ISiteInfoWebPartProps {
  description: string;
}

export default class SiteInfoWebPart extends BaseClientSideWebPart<ISiteInfoWebPartProps> {
  private el: HTMLDivElement = null;

  // Method to render the webpart
  public render(): void {
    // Set the context
    $REST.ContextInfo.setPageContext(this.context.pageContext);

    // Set the html template
    this.domElement.innerHTML = `
      <div class="${ styles.siteInfo}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div id="site-info" class="fabric ${ styles.column}">
            </div>
          </div>
        </div>
      </div>`;

    // Get the site info element
    this.el = this.domElement.querySelector("#site-info") as HTMLDivElement;

    // Load the information
    this.load();
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
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Private Methods
   */

  // Loads the current web information
  private load(url?: string) {
    // Render a spinner
    Fabric.Spinner({
      el: this.el,
      text: "Loading the Site Information"
    });

    // Get the web information
    $REST.Web(url).query({
      Expand: ["ContentTypes", "Fields", "Lists", "Webs"]
    }).execute(web => {
      // Render the tabs
      Fabric.Pivot({
        el: this.el,
        tabs: [
          {
            isSelected: true,
            name: "Sub Webs",
            content: this.renderSubWebs(web.Webs.results)
          },
          {
            name: "Content Types",
            content: this.renderContentTypes(web.ContentTypes.results)
          },
          {
            name: "Fields",
            content: this.renderFields(web.Fields.results)
          },
          {
            name: "Lists",
            content: this.renderLists(web.Lists.results)
          },
        ]
      });
    });
  }

  // Renders the content types
  private renderContentTypes(contentTypes: Array<Types.SP.IContentTypeResult>) {
    let items = [];

    // Sort the content types
    contentTypes = contentTypes.sort((a, b) => {
      if (a.Name < b.Name) { return -1; }
      if (a.Name > b.Name) { return 1; }
      return 0;
    });

    // Parse the content types
    for (let i = 0; i < contentTypes.length; i++) {
      let contentType = contentTypes[i];

      // Add the item
      items.push(Fabric.Templates.ListItem({
        primaryText: contentType.Name,
        secondaryText: contentType.Description
      }));
    }

    // Render a list
    return Fabric.Templates.List({ items });
  }

  // Renders the fields
  private renderFields(fields: Array<Types.SP.IFieldResult>) {
    let items = [];

    // Sort the fields
    fields = fields.sort((a, b) => {
      if (a.Title < b.Title) { return -1; }
      if (a.Title > b.Title) { return 1; }
      return 0;
    });

    // Parse the fields
    for (let i = 0; i < fields.length; i++) {
      let field = fields[i];

      // Add the item
      items.push(Fabric.Templates.ListItem({
        primaryText: field.Title,
        secondaryText: field.Description,
        tertiaryText: field.TypeAsString,
        metaText: field.InternalName
      }));
    }

    // Render a list
    return Fabric.Templates.List({ items });
  }

  // Renders the lists
  private renderLists(lists: Array<Types.SP.IListResult>) {
    let items = [];

    // Sort the lists
    lists = lists.sort((a, b) => {
      if (a.Title < b.Title) { return -1; }
      if (a.Title > b.Title) { return 1; }
      return 0;
    });

    // Parse the lists
    for (let i = 0; i < lists.length; i++) {
      let list = lists[i];

      // Add the item
      items.push(Fabric.Templates.ListItem({
        primaryText: list.Title,
        secondaryText: list.Description,
        metaText: list.BaseTemplate + ""
      }));
    }

    // Render a list
    return Fabric.Templates.List({ items });
  }

  // Renders the sub webs
  private renderSubWebs(webs: Array<Types.SP.IWebResult>) {
    let items = [];

    // Sort the webs
    webs = webs.sort((a, b) => {
      if (a.Title < b.Title) { return -1; }
      if (a.Title > b.Title) { return 1; }
      return 0;
    });

    // Parse the webs
    for (let i = 0; i < webs.length; i++) {
      let web = webs[i];

      // Add the item
      items.push(Fabric.Templates.ListItem({
        primaryText: web.Title,
        secondaryText: web.Description,
        metaText: web.ServerRelativeUrl
      }));
    }

    // Render a list
    return Fabric.Templates.List({ items });
  }
}