# SharePoint Custom CSS Extension

## Summary

This SharePoint Framework (SPFx) extension allows you to apply custom CSS to specific SharePoint Online pages based on URL patterns. It provides a flexible way to customize the appearance of your SharePoint site without modifying the core templates.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Solution

| Solution               | Author(s)                             |
| ---------------------- | ------------------------------------- |
| sp-custom-extension    | Your Name                             |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | April 25, 2025   | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - **npm install**
  - **gulp build**
  - **gulp bundle --ship**
  - **gulp package-solution --ship**
- Upload the .sppkg file from the `sharepoint/solution` folder to your App Catalog
- Add the app to your site
- Configure the extension properties as needed

## Features

This extension allows you to customize the appearance of SharePoint pages by applying custom CSS styles with URL-based conditional logic. The extension:

- Loads custom CSS only on specific pages that match URL patterns
- Supports both inclusion and exclusion patterns for precise control
- Allows for an optional additional CSS file to be loaded from a URL specified in the properties
- Includes debug mode for troubleshooting URL pattern matching

### Configuration Options

The extension supports the following configuration options:

| Property | Type | Description |
| -------- | ---- | ----------- |
| useInternalCss | boolean | Whether to use the built-in CSS file (default: true) |
| cssUrl | string | URL to an external CSS file to load (optional) |
| includeUrls | string[] | Array of URL patterns where CSS should be applied |
| excludeUrls | string[] | Array of URL patterns where CSS should NOT be applied |
| enableDebug | boolean | Enable logging of URL matching to browser console |

### How URL Matching Works

- The `includeUrls` array contains URL patterns that, if matched, will have the CSS applied
- The `excludeUrls` array contains URL patterns that, if matched, will NOT have the CSS applied
- Exclude patterns take precedence over include patterns
- If no include patterns are specified, CSS will be applied everywhere (except excluded patterns)
- If include patterns are specified, CSS will ONLY be applied on pages matching those patterns
- The matching is case-insensitive and uses simple substring matching

### Configuration Examples

#### Example 1: Apply CSS only on specific pages

```json
"properties": {
  "includeUrls": [
    "/SitePages/Project-A.aspx",
    "/SitePages/Project-B.aspx",
    "/SitePages/Dashboard.aspx"
  ]
}
```

#### Example 2: Apply CSS on all pages except specific ones

```json
"properties": {
  "excludeUrls": [
    "/SitePages/Home.aspx",
    "/SitePages/Contact.aspx",
    "/Lists/"
  ]
}
```

#### Example 3: Apply CSS only on certain sections of your site

```json
"properties": {
  "includeUrls": [
    "/SitePages/Projects/",
    "/SitePages/Reports/"
  ]
}
```

#### Example 4: Use an external CSS file instead of the built-in one

```json
"properties": {
  "useInternalCss": false,
  "cssUrl": "https://yoursite.sharepoint.com/sites/yoursite/SiteAssets/custom.css"
}
```

## Customizing the CSS

The built-in CSS file is located at `src/extensions/spCustomCssExtension/styles/customStyles.css`. Modify this file to customize the appearance of your SharePoint pages.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
