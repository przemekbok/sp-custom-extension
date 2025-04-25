import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'SpCustomCssExtensionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpCustomCssExtensionApplicationCustomizer';

/**
 * Interface for properties of the custom CSS extension
 */
export interface ISpCustomCssExtensionApplicationCustomizerProperties {
  // Path to an optional external CSS file
  cssUrl?: string;
  // Whether to use the internal CSS (default: true)
  useInternalCss?: boolean;
  // Array of URL patterns to include (apply CSS only on these pages)
  includeUrls?: string[];
  // Array of URL patterns to exclude (don't apply CSS on these pages)
  excludeUrls?: string[];
  // Whether to log debug information
  enableDebug?: boolean;
}

/** A Custom Action which injects custom CSS into SharePoint pages */
export default class SpCustomCssExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpCustomCssExtensionApplicationCustomizerProperties> {

  private _internalCssPath: string = require('./styles/customStyles.css');

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    // Get current page URL
    const currentUrl = window.location.href.toLowerCase();
    
    // Debug logging if enabled
    if (this.properties.enableDebug) {
      console.log(`[${LOG_SOURCE}] Current URL: ${currentUrl}`);
      console.log(`[${LOG_SOURCE}] Include patterns:`, this.properties.includeUrls);
      console.log(`[${LOG_SOURCE}] Exclude patterns:`, this.properties.excludeUrls);
    }
    
    // Check if page should be excluded
    if (this.properties.excludeUrls && this.properties.excludeUrls.length > 0) {
      for (const pattern of this.properties.excludeUrls) {
        if (currentUrl.indexOf(pattern.toLowerCase()) > -1) {
          if (this.properties.enableDebug) {
            console.log(`[${LOG_SOURCE}] URL matched exclude pattern "${pattern}" - CSS will not be applied`);
          }
          return Promise.resolve(); // Exit without applying CSS
        }
      }
    }
    
    // Check if page should be included (only if includeUrls is specified)
    let shouldApplyCss = true;
    
    if (this.properties.includeUrls && this.properties.includeUrls.length > 0) {
      shouldApplyCss = false; // Default to false when include patterns exist
      
      for (const pattern of this.properties.includeUrls) {
        if (currentUrl.indexOf(pattern.toLowerCase()) > -1) {
          shouldApplyCss = true;
          if (this.properties.enableDebug) {
            console.log(`[${LOG_SOURCE}] URL matched include pattern "${pattern}" - CSS will be applied`);
          }
          break;
        }
      }
    }
    
    // Apply CSS if conditions are met
    if (shouldApplyCss) {
      // Load internal CSS if not explicitly disabled
      const useInternalCss = this.properties.useInternalCss !== false;
      if (useInternalCss) {
        SPComponentLoader.loadCss(this._internalCssPath);
        Log.info(LOG_SOURCE, `Loaded internal CSS file`);
      }

      // Load external CSS if provided
      if (this.properties.cssUrl) {
        SPComponentLoader.loadCss(this.properties.cssUrl);
        Log.info(LOG_SOURCE, `Loaded external CSS from: ${this.properties.cssUrl}`);
      }
    } else if (this.properties.enableDebug) {
      console.log(`[${LOG_SOURCE}] No include patterns matched - CSS will not be applied`);
    }

    return Promise.resolve();
  }
}