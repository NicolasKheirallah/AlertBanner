export class SiteIdUtils {
  public static normalizeGuid(guid: string): string {
    if (!guid) return "";
    return guid.replace(/[{}]/g, "").toLowerCase();
  }

  public static extractGuidFromGraphId(graphId: string): string | null {
    if (!graphId) return null;
    
    if (graphId.includes(",")) {
      const parts = graphId.split(",");
      if (parts.length >= 2) {
        return this.normalizeGuid(parts[1]);
      }
    }
    
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

    if (normalized.includes(",")) {
      const parts = normalized.split(",");
      if (parts.length >= 2) {
        variations.add(this.normalizeGuid(parts[1])); // Site GUID
        if (parts[0]) variations.add(parts[0].trim()); // Hostname
      }
    }

    // Check if it looks like a URL (contains /) or hostname (contains .)
    if (normalized.includes("/") || normalized.includes(".")) {
       try {
         const urlStr = normalized.startsWith("http") ? normalized : `https://${normalized}`;
         const url = new URL(urlStr);
         
         variations.add(url.hostname.toLowerCase());
         
         const cleanPath = url.pathname.endsWith('/') && url.pathname.length > 1 
            ? url.pathname.slice(0, -1) 
            : url.pathname;
            
         variations.add((url.hostname + cleanPath).toLowerCase());
       } catch (e) {
       }
    }
    
    return Array.from(variations);
  }
}
