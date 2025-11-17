import { CommandArguments, File, Folder } from "@models";
import {
  execScript,
  CliCommand,
  ListHelpers,
  Logger,
  ArgumentsHelper,
} from "@helpers";
import { basename } from "path";

const normalizeValue = (value?: string | null) =>
  (value || "").toString().toLowerCase();

export class FileHelpers {
  private static allPages: File[] = [];
  private static checkedFiles: string[] = [];

  /**
   * Retrieve the relative path for the file
   * @param webUrl
   * @param library
   * @param filePath
   */
  public static getRelUrl(webUrl: string, filePath: string) {
    const relWebUrl = webUrl.split("sharepoint.com").pop();
    return `${relWebUrl.startsWith("/") ? "" : "/"}${relWebUrl}${
      relWebUrl.endsWith("/") ? "" : "/"
    }${filePath}`;
  }

  /**
   * Create the file on SharePoint
   * @param crntFolder
   * @param imgPath
   * @param webUrl
   * @param override
   */
  public static async create(
    crntFolder: string,
    imgPath: string,
    webUrl: string,
    override: boolean = false
  ) {
    Logger.debug(`Create file "${imgPath}" to "${crntFolder}"`);
    const cacheKey = `${imgPath.replace(/ /g, "%20")}-${crntFolder.replace(
      / /g,
      "%20"
    )}`;
    if (this.checkedFiles && this.checkedFiles.indexOf(cacheKey) === -1) {
      if (override) {
        await this.upload(webUrl, crntFolder, imgPath);
      } else {
        try {
          // Check if file exists
          const filePath = `${crntFolder}/${basename(imgPath)}`;
          const relativeUrl = this.getRelUrl(webUrl, filePath);
          const fileData = await execScript(
            ArgumentsHelper.parse(
              `spo file get --webUrl "${webUrl}" --url "${relativeUrl}"`
            ),
            false
          );
          Logger.debug(`File data retrieved: ${JSON.stringify(fileData)}`);
        } catch (e) {
          await this.upload(webUrl, crntFolder, imgPath);
        }
      }

      this.checkedFiles.push(cacheKey);
    }

    return `${webUrl}/${crntFolder}/${basename(imgPath)}`.replace(/ /g, "%20");
  }

  /**
   * Clean up all files in the folder
   * @param options
   */
  public static async cleanUp(options: CommandArguments, crntFolder: string) {
    if (options.cleanStart) {
      try {
        const { webUrl } = options;
        let filesData: File[] | string = await execScript<string>(
          ArgumentsHelper.parse(
            `spo file list --webUrl "${webUrl}" -f "${crntFolder}" -o json`
          ),
          CliCommand.getRetry()
        );
        if (filesData && typeof filesData === "string") {
          filesData = JSON.parse(filesData);
        }

        Logger.debug(`Files to be removed: ${JSON.stringify(filesData)}`);

        for (const file of filesData as File[]) {
          if (!file || !file.ServerRelativeUrl) {
            continue;
          }
          const normalizedServerUrl = normalizeValue(file.ServerRelativeUrl);
          if (!normalizedServerUrl) {
            continue;
          }
          const normalizedFolder = normalizeValue(crntFolder);
          const relPath =
            normalizedServerUrl.split(normalizedFolder).pop() || "";
          const filePath = `${crntFolder}${relPath}`;
          await execScript<string>(
            ArgumentsHelper.parse(
              `spo file remove --webUrl "${webUrl}" --url "${filePath}" --force`
            ),
            CliCommand.getRetry()
          );
        }

        let folderData: Folder[] | string = await execScript<string>(
          ArgumentsHelper.parse(
            `spo folder list --webUrl "${webUrl}" --parentFolderUrl "${crntFolder}" -o json`
          ),
          CliCommand.getRetry()
        );
        if (folderData && typeof folderData === "string") {
          folderData = JSON.parse(folderData);
        }

        Logger.debug(`Folders to be removed: ${JSON.stringify(folderData)}`);

        for (const folder of folderData as Folder[]) {
          if (!folder || !folder.Exists) {
            continue;
          }
          const folderName = normalizeValue(folder.Name);
          if (!folderName || folderName === "forms" || folderName === "templates") {
            continue;
          }
          const normalizedServerUrl = normalizeValue(folder.ServerRelativeUrl);
          if (!normalizedServerUrl) {
            continue;
          }
          const normalizedFolder = normalizeValue(crntFolder);
          const relPath =
            normalizedServerUrl.split(normalizedFolder).pop() || "";
          const folderPath = `${crntFolder}${relPath}`;
          await execScript<string>(
            ArgumentsHelper.parse(
              `spo folder remove --webUrl "${webUrl}" --url "${folderPath}" --force`
            ),
            CliCommand.getRetry()
          );
        }
      } catch (err) {
        throw (err instanceof Error ? err : new Error(err as any));
      }
    }
  }

  /**
   * Retrieve all pages
   * @param webUrl
   * @param crntFolder
   */
  public static async getAllPages(
    webUrl: string,
    crntFolder: string
  ): Promise<File[]> {
    if (this.allPages && this.allPages.length > 0) {
      return this.allPages;
    }

    const pageList = await ListHelpers.getSitePagesList(webUrl);

    let filesData: File[] | string = await execScript<string>(
      ArgumentsHelper.parse(
        `spo listitem list --webUrl "${webUrl}" --id "${pageList.Id}" --fields "ID,Title,FileRef" -o json`
      ),
      CliCommand.getRetry()
    );
    if (filesData && typeof filesData === "string") {
      filesData = JSON.parse(filesData);
    }

    this.allPages = filesData as File[];
    return this.allPages;
  }

  /**
   * Upload the file
   * @param webUrl
   * @param crntFolder
   * @param imgPath
   */
  private static async upload(
    webUrl: string,
    crntFolder: string,
    imgPath: string
  ) {
    Logger.debug(`Uploading file "${imgPath}" to ${crntFolder}"`);
    await execScript(
      ArgumentsHelper.parse(
        `spo file add --webUrl "${webUrl}" --folder "${crntFolder}" --path "${imgPath}"`
      ),
      CliCommand.getRetry()
    );
  }
}
