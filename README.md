## spfx-siteinfo
![demo](https://github.com/gunjandatta/spfx-siteinfo/raw/master/images/demo.png)
This is a SPFX modern webpart displaying the web information. This example uses the [gd-sprest-js](https://gunjandatta.github.io/js) library.

### Building the code

```bash
gulp serve
```

### Code Examples
Below are the code examples from this demo that are worth highlighting.
#### Import Library
We will be using the gd-sprest and gd-sprest-js libraries for interacting with the REST API and rendering components using the Office Fabric-UI JavaScript framework.

```ts
import { $REST, Types } from "gd-sprest";
import { Fabric } from "gd-sprest-js";
```

#### Reference Page Context

The gd-sprest library will require you to set the page context for it to work in modern pages.

```ts
$REST.ContextInfo.setPageContext(this.context.pageContext);
```

#### Fabric Spinner

```ts
// Render a loading message
Fabric.Spinner({
  el: this.domElement.querySelector("#site-info"),
  text: "Loading the Site Information"
});

```

#### Web Query

This demo will execute 1 request to the server for the information.

```ts
// Get the current web
$REST.Web()
    // Set the query
    .query({
      Expand: ["ContentTypes", "Fields", "Lists", "Webs"]
    })
    // Execute the request
    .execute(web => {
      // Render the tabs
      renderTabs(web);
    });
```

#### Fabric Pivot

The gd-sprest-js "Fabric" class has a "Templates" class that can be used to render the raw html. We will use this to render the html for the "Lists" component.

```ts
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
```