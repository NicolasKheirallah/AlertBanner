export class SiteIdUtils {
  public static normalizeGuid(guid: string): string {
    if (!guid) return "";
    return guid.replace(/[{}]/g, "").toLowerCase();
  }

  public static extractGuidFromGraphId(graphId: string): string | null {
    if (!graphId) return null;
    
    // deeply composite id matching pattern
    if (graphId.includes(",")) {
      const parts = graphId.split(",");
      // Usually parts[1] is the site GUID in standard Graph ID
      if (parts.length >= 2) {
        return this.normalizeGuid(parts[1]);
      }
    }
    
    // Fallback: assume it might be a simple GUID
    if (this.isGuid(graphId)) {
        return this.normalizeGuid(graphId);
    }

    return null;
  }

  public static isGuid(str: string): boolean {
    const guidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    return guidRegex.test(this.normalizeGuid(str));
  }

  public static createDedupKey(siteId: string): string {
    if (!siteId) return "";
    
    if (siteId.includes(",")) {
      const guid = this.extractGuidFromGraphId(siteId);
      if (guid) return guid;
    }
    
    if (siteId.includes(":")) {
        return siteId.toLowerCase();
    }

    return this.normalizeGuid(siteId);
  }

  public static generateSiteVariations(input: string): string[] {
    if (!input) return [];

    const variations: Set<string> = new Set();
    const normalized = input.toLowerCase().trim();
    
    variations.add(normalized);

    // Handle Composite Graph ID
    if (normalized.includes(",")) {
      const parts = normalized.split(",");
      // Format: hostname,siteGuid,webGuid
      if (parts.length >= 2) {
        variations.add(this.normalizeGuid(parts[1])); // Site GUID
        if (parts[0]) variations.add(parts[0].trim()); // Hostname
      }
    }

    // Handle URL or Hostname
    // Check if it looks like a URL (contains /) or hostname (contains .)
    if (normalized.includes("/") || normalized.includes(".")) {
       try {
         // ensuring protocol for URL parsing
         const urlStr = normalized.startsWith("http") ? normalized : `https://${normalized}`;
         const url = new URL(urlStr);
         
         // Add hostname (e.g. contoso.sharepoint.com)
         variations.add(url.hostname.toLowerCase());
         
         // Add hostname + pathname (e.g. contoso.sharepoint.com/sites/hr)
         // Remove trailing slash if present for consistency
         const cleanPath = url.pathname.endsWith('/') && url.pathname.length > 1 
            ? url.pathname.slice(0, -1) 
            : url.pathname;
            
         variations.add((url.hostname + cleanPath).toLowerCase());
       } catch (e) {
         // ignore parse errors
       }
    }
    
    return Array.from(variations);
  }
}
