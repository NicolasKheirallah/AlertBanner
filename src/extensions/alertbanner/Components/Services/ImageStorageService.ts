import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClientV3,
} from "@microsoft/sp-http";
import { logger } from "./LoggerService";
import { SharePointListLocator } from "./SharePointListLocator";

export interface IExistingImage {
  name: string;
  serverRelativeUrl: string;
  timeCreated: string;
  length: number;
}

const ALERT_IMAGES_FOLDER = "AlertBannerImages";

export class ImageStorageService {
  private graphSiteId?: string;
  private graphDriveId?: string;
  private graphClient?: MSGraphClientV3;
  private ensureDrivePromise?: Promise<void>;
  private locator?: SharePointListLocator;

  constructor(private readonly context: ApplicationCustomizerContext) {}

  public async uploadImage(file: File, folderName?: string): Promise<string> {
    const sanitizedFolder = folderName
      ? this.sanitizeFolderName(folderName)
      : undefined;
    const uniqueFileName = this.getUniqueFileName(file);

    try {
      return await this.uploadViaGraph(file, uniqueFileName, sanitizedFolder);
    } catch (graphError) {
      logger.warn(
        "ImageStorageService",
        "Graph upload failed; falling back to SharePoint REST",
        graphError,
      );
      return await this.uploadViaRest(file, uniqueFileName, sanitizedFolder);
    }
  }

  public async listImages(
    folderName?: string,
    siteId?: string,
  ): Promise<IExistingImage[]> {
    if (!folderName) {
      return [];
    }

    const sanitizedFolder = this.sanitizeFolderName(folderName);
    const { siteUrl, siteAssetsRoot } = await this.getSitePaths(siteId);
    const folderPath = `${siteAssetsRoot}/${ALERT_IMAGES_FOLDER}/${sanitizedFolder}`;
    const normalizedFolder = folderPath.startsWith("/")
      ? folderPath
      : `/${folderPath}`;
    const escapedFolder = normalizedFolder.replace(/'/g, "''");
    const filesUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolder}')/Files?$select=Name,ServerRelativeUrl,TimeCreated,Length`;

    try {
      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          filesUrl,
          SPHttpClient.configurations.v1,
        );

      if (response.status === 404) {
        return [];
      }

      if (!response.ok) {
        const message = await response.text();
        throw new Error(`Failed to load images: ${response.status} ${message}`);
      }

      const data = await response.json();
      const files: {
        id: string;
        Name: string;
        ServerRelativeUrl: string;
        TimeCreated: string;
        Length: string | number;
        "@microsoft.graph.downloadUrl"?: string;
        file?: unknown;
        folder?: unknown;
      }[] = data?.value ?? [];

      return files
        .filter((file) => /\.(jpg|jpeg|png|gif|webp)$/i.test(file.Name))
        .map((file) => ({
          name: file.Name,
          serverRelativeUrl: file.ServerRelativeUrl,
          timeCreated: file.TimeCreated,
          length:
            typeof file.Length === "string"
              ? parseInt(file.Length, 10)
              : file.Length || 0,
        }));
    } catch (error) {
      logger.warn("ImageStorageService", "Failed to list images", error);
      throw error;
    }
  }

  public async deleteImage(
    fileName: string,
    folderName: string,
    siteId?: string,
  ): Promise<void> {
    const sanitizedFolder = this.sanitizeFolderName(folderName);
    const { siteUrl, siteAssetsRoot } = await this.getSitePaths(siteId);
    const folderPath = `${siteAssetsRoot}/${ALERT_IMAGES_FOLDER}/${sanitizedFolder}`;
    const normalizedFolder = folderPath.startsWith("/")
      ? folderPath
      : `/${folderPath}`;

    const serverRelativeUrl = `${normalizedFolder}/${fileName}`;
    const escapedFileUrl = serverRelativeUrl.replace(/'/g, "''");

    const deleteUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${escapedFileUrl}')`;

    try {
      const response = await this.context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*",
          },
        },
      );

      if (!response.ok) {
        throw new Error(`Failed to delete image: ${response.statusText}`);
      }

      logger.info("ImageStorageService", "Deleted image", {
        fileName,
        folderName,
        siteId,
      });
    } catch (error) {
      logger.error("ImageStorageService", "Error deleting image", error);
      throw error;
    }
  }

  public async deleteImageFolder(
    folderName: string,
    siteId?: string,
  ): Promise<void> {
    const sanitizedFolder = this.sanitizeFolderName(folderName);
    const { siteUrl, siteAssetsRoot } = await this.getSitePaths(siteId);
    const folderPath = `${siteAssetsRoot}/${ALERT_IMAGES_FOLDER}/${sanitizedFolder}`;
    const normalizedFolder = folderPath.startsWith("/")
      ? folderPath
      : `/${folderPath}`;
    const escapedFolder = normalizedFolder.replace(/'/g, "''");

    const deleteFolderUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolder}')`;

    try {
      await this.context.spHttpClient.post(
        deleteFolderUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*",
          },
        },
      );
      logger.info("ImageStorageService", "Image folder deleted successfully", {
        folderName,
        siteId,
      });
    } catch (error) {
      // Folder deletion is optional - log warning if it fails
      logger.warn(
        "ImageStorageService",
        "Could not delete image folder (may not exist)",
        { folderName, siteId, error },
      );
    }
  }

  private sanitizeFolderName(name: string): string {
    let sanitized = name
      .replace(/[\u{1F300}-\u{1F9FF}]/gu, "")
      .replace(/[\u{2600}-\u{26FF}]/gu, "")
      .replace(/[\u{2700}-\u{27BF}]/gu, "")
      .replace(/[\u{1F000}-\u{1F2FF}]/gu, "")
      .replace(/[\u{1FA00}-\u{1FAFF}]/gu, "")
      .replace(/[\u{FE00}-\u{FEFF}]/gu, "")
      .replace(/[^\x00-\x7F]/g, "")
      .replace(/[~#%&*{}<>?/|":\\]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^[_.]+|[_.]+$/g, "")
      .trim();

    if (sanitized.length === 0) {
      sanitized = "AlertFolder_" + Date.now();
    }

    return sanitized.substring(0, 128);
  }

  private getUniqueFileName(file: File): string {
    const extension = file.name.includes(".")
      ? file.name.substring(file.name.lastIndexOf("."))
      : "";
    const rawBaseName = file.name.includes(".")
      ? file.name.substring(0, file.name.lastIndexOf("."))
      : file.name;
    let sanitizedBase = rawBaseName
      .replace(/[^a-zA-Z0-9]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");

    if (!sanitizedBase) {
      sanitizedBase = "Image";
    }

    return `${sanitizedBase}_${Date.now()}${extension}`;
  }

  private async uploadViaGraph(
    file: File,
    uniqueFileName: string,
    sanitizedFolder?: string,
  ): Promise<string> {
    const { graphClient, driveId } = await this.ensureGraphContext();

    const folderSegments = sanitizedFolder
      ? [ALERT_IMAGES_FOLDER, sanitizedFolder]
      : [ALERT_IMAGES_FOLDER];
    const targetFolderId = await this.ensureGraphFolders(
      driveId,
      folderSegments,
      graphClient,
    );

    const tokenProvider =
      await this.context.aadTokenProviderFactory.getTokenProvider();
    const token = await tokenProvider.getToken("https://graph.microsoft.com");
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${targetFolderId}:/${encodeURIComponent(uniqueFileName)}:/content`;

    const uploadResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": file.type || "application/octet-stream",
      },
      body: file,
    });

    if (!uploadResponse.ok) {
      const message = await uploadResponse.text();
      throw new Error(
        `Microsoft Graph upload failed: ${uploadResponse.status} ${message}`,
      );
    }

    const uploaded = await uploadResponse.json();
    const uploadedUrl = uploaded?.webUrl as string | undefined;

    if (!uploadedUrl) {
      throw new Error("Microsoft Graph upload did not return a webUrl.");
    }

    return uploadedUrl;
  }

  private async uploadViaRest(
    file: File,
    uniqueFileName: string,
    sanitizedFolder?: string,
  ): Promise<string> {
    const { siteUrl, siteAssetsRoot } = await this.getSitePaths();
    const baseFolderPath = `${siteAssetsRoot}/${ALERT_IMAGES_FOLDER}`;

    await this.ensureFolderExistsRest(siteUrl, siteAssetsRoot);
    await this.ensureFolderExistsRest(siteUrl, baseFolderPath);

    let targetFolderPath = baseFolderPath;
    if (sanitizedFolder) {
      targetFolderPath = `${baseFolderPath}/${sanitizedFolder}`;
      await this.ensureFolderExistsRest(siteUrl, targetFolderPath);
    }

    const safeTargetFolder = targetFolderPath.startsWith("/")
      ? targetFolderPath
      : `/${targetFolderPath}`;
    const encodedFolderForApi = encodeURIComponent(safeTargetFolder);
    const uniqueFileNameForApi = encodeURIComponent(uniqueFileName);

    const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedFolderForApi}')/Files/add(url='${uniqueFileNameForApi}',overwrite=true)`;

    const uploadHeaders: Record<string, string> = {
      Accept: "application/json;odata=verbose",
      "Content-Type": file.type || "application/octet-stream",
      binaryStringRequestBody: "true",
    };

    const digestValue = (
      this.context.pageContext as unknown as {
        legacyPageContext?: { formDigestValue?: string };
      }
    )?.legacyPageContext?.formDigestValue;
    if (digestValue) {
      uploadHeaders["X-RequestDigest"] = digestValue;
    }

    const uploadResponse = await this.context.spHttpClient.post(
      uploadUrl,
      SPHttpClient.configurations.v1,
      {
        headers: uploadHeaders,
        body: file,
      },
    );

    if (!uploadResponse.ok) {
      const message = await uploadResponse.text();
      throw new Error(`Upload failed: ${uploadResponse.status} ${message}`);
    }

    const serverRelativeFileUrl = `${safeTargetFolder.replace(/\/$/, "")}/${uniqueFileName}`;
    const fullImageUrl = new URL(
      serverRelativeFileUrl,
      this.context.pageContext.web.absoluteUrl,
    ).toString();

    return fullImageUrl;
  }

  private async ensureGraphContext(): Promise<{
    graphClient: MSGraphClientV3;
    driveId: string;
    siteId: string;
  }> {
    if (this.graphClient && this.graphDriveId && this.graphSiteId) {
      return {
        graphClient: this.graphClient,
        driveId: this.graphDriveId,
        siteId: this.graphSiteId,
      };
    }

    if (!this.ensureDrivePromise) {
      this.ensureDrivePromise = (async () => {
        this.graphClient =
          await this.context.msGraphClientFactory.getClient("3");

        const currentUrl = new URL(this.context.pageContext.web.absoluteUrl);
        const hostname = currentUrl.hostname;
        let sitePath = currentUrl.pathname || "/";
        if (!sitePath.endsWith("/")) {
          sitePath = `${sitePath}/`;
        }

        const siteInfo = await this.graphClient!.api(
          `/sites/${hostname}:${sitePath}`,
        )
          .select("id")
          .get();

        const siteId = siteInfo?.id;
        if (!siteId) {
          throw new Error("Unable to resolve site ID via Microsoft Graph.");
        }

        const drivesResponse = await this.graphClient!.api(
          `/sites/${siteId}/drives`,
        )
          .filter("name eq 'SiteAssets'")
          .get();

        const siteAssetsDrive = drivesResponse?.value?.[0];
        if (!siteAssetsDrive?.id) {
          throw new Error("SiteAssets drive not found via Microsoft Graph.");
        }

        this.graphSiteId = siteId;
        this.graphDriveId = siteAssetsDrive.id;
      })();
    }

    await this.ensureDrivePromise;
    return {
      graphClient: this.graphClient!,
      driveId: this.graphDriveId!,
      siteId: this.graphSiteId!,
    };
  }

  private async ensureGraphFolders(
    driveId: string,
    segments: string[],
    graphClient: MSGraphClientV3,
  ): Promise<string> {
    let parentId: string | undefined;

    for (const segment of segments) {
      if (!parentId) {
        parentId = await this.ensureGraphRootChild(
          driveId,
          segment,
          graphClient,
        );
      } else {
        parentId = await this.ensureGraphChild(
          driveId,
          parentId,
          segment,
          graphClient,
        );
      }
    }

    return parentId!;
  }

  private async ensureGraphRootChild(
    driveId: string,
    folderName: string,
    graphClient: MSGraphClientV3,
  ): Promise<string> {
    try {
      const existing = await graphClient
        .api(`/drives/${driveId}/root:/${folderName}`)
        .get();
      return existing.id;
    } catch (error: unknown) {
      if (this.isMissing(error)) {
        const created = await graphClient
          .api(`/drives/${driveId}/root/children`)
          .post({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
          });
        return created.id;
      }

      if (this.isConflict(error)) {
        const existing = await graphClient
          .api(`/drives/${driveId}/root:/${folderName}`)
          .get();
        return existing.id;
      }

      throw error;
    }
  }

  private async ensureGraphChild(
    driveId: string,
    parentId: string,
    folderName: string,
    graphClient: MSGraphClientV3,
  ): Promise<string> {
    try {
      const existing = await graphClient
        .api(`/drives/${driveId}/items/${parentId}:/${folderName}`)
        .get();
      return existing.id;
    } catch (error: unknown) {
      if (this.isMissing(error)) {
        const created = await graphClient
          .api(`/drives/${driveId}/items/${parentId}/children`)
          .post({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
          });
        return created.id;
      }

      if (this.isConflict(error)) {
        const children = await graphClient
          .api(`/drives/${driveId}/items/${parentId}/children`)
          .get();
        const match = children?.value?.find(
          (item: { name: string }) => item.name === folderName,
        );
        if (match?.id) {
          return match.id;
        }
      }

      throw error;
    }
  }

  private async ensureFolderExistsRest(
    siteUrl: string,
    folderServerRelativeUrl: string,
  ): Promise<void> {
    const normalizedFolder = folderServerRelativeUrl.startsWith("/")
      ? folderServerRelativeUrl
      : `/${folderServerRelativeUrl}`;

    const escapedFolder = normalizedFolder.replace(/'/g, "''");
    const getFolderEndpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedFolder}')`;

    const getResponse = await this.context.spHttpClient.get(
      getFolderEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "OData-Version": "3.0",
        },
      },
    );

    if (getResponse.ok) {
      return;
    }

    if (getResponse.status === 404) {
      const inheritResponse = await this.context.spHttpClient.post(
        `${siteUrl}/_api/contextinfo`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "OData-Version": "3.0",
          },
        },
      );

      if (!inheritResponse.ok) {
        const message = await inheritResponse.text();
        throw new Error(
          `Failed to obtain form digest: ${inheritResponse.status} ${message}`,
        );
      }

      const contextInfo = await inheritResponse.json();
      const formDigestValue =
        contextInfo?.FormDigestValue ??
        contextInfo?.d?.GetContextWebInformation?.FormDigestValue ??
        contextInfo?.GetContextWebInformation?.FormDigestValue;

      if (!formDigestValue) {
        throw new Error(
          "Failed to resolve form digest value for folder creation.",
        );
      }

      const createResponse = await this.context.spHttpClient.post(
        `${siteUrl}/_api/web/Folders`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "OData-Version": "3.0",
            "X-RequestDigest": formDigestValue,
          },
          body: JSON.stringify({
            __metadata: { type: "SP.Folder" },
            ServerRelativeUrl: normalizedFolder,
          }),
        },
      );

      if (!createResponse.ok && createResponse.status !== 409) {
        const message = await createResponse.text();
        throw new Error(
          `Failed to create folder ${normalizedFolder}: ${createResponse.status} ${message}`,
        );
      }

      if (createResponse.ok) {
        logger.info("ImageStorageService", "Created folder via REST", {
          folder: normalizedFolder,
        });
      }
    } else {
      const message = await getResponse.text();
      throw new Error(
        `Failed to access folder ${normalizedFolder}: ${getResponse.status} ${message}`,
      );
    }
  }

  private async getSitePaths(
    siteId?: string,
  ): Promise<{ siteUrl: string; siteAssetsRoot: string }> {
    let siteUrl = this.context.pageContext.web.absoluteUrl;

    if (siteId) {
      const currentSiteGuid = this.context.pageContext.site.id
        .toString()
        .replace(/[{}]/g, "")
        .toLowerCase();
      const requestedSiteGuid = siteId.replace(/[{}]/g, "").toLowerCase();
      if (requestedSiteGuid && requestedSiteGuid !== currentSiteGuid) {
        try {
          const locator = await this.ensureLocator();
          siteUrl = await locator.getSiteUrlFromIdentifier(siteId);
        } catch (error) {
          logger.warn(
            "ImageStorageService",
            "Failed to resolve site URL for image operation; using current site",
            {
              siteId,
              error,
            },
          );
        }
      }
    }

    const normalizedSiteUrl = siteUrl.replace(/\/$/, "");
    const parsed = new URL(normalizedSiteUrl);
    const rawPath = parsed.pathname || "";
    const trimmedServerRelativeWebUrl =
      rawPath === "/" ? "" : rawPath.replace(/\/$/, "");
    const siteAssetsRoot = `${trimmedServerRelativeWebUrl ? trimmedServerRelativeWebUrl : ""}/SiteAssets`;

    return { siteUrl: normalizedSiteUrl, siteAssetsRoot };
  }

  private async ensureLocator(): Promise<SharePointListLocator> {
    if (this.locator) {
      return this.locator;
    }

    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient("3");
    }

    this.locator = new SharePointListLocator(this.graphClient, this.context);
    return this.locator;
  }

  private isMissing(error: unknown): boolean {
    const err = error as {
      statusCode?: number;
      status?: number;
      code?: string;
      body?: { error?: { code?: string } };
    } | null;
    const statusCode = err?.statusCode ?? err?.status;
    const code = err?.code ?? err?.body?.error?.code;
    return statusCode === 404 || code === "itemNotFound";
  }

  private isConflict(error: unknown): boolean {
    const err = error as {
      statusCode?: number;
      status?: number;
      code?: string;
      body?: { error?: { code?: string } };
    } | null;
    const statusCode = err?.statusCode ?? err?.status;
    const code = err?.code ?? err?.body?.error?.code;
    return statusCode === 409 || code === "nameAlreadyExists";
  }
}
