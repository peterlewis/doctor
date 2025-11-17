import { Observable } from "rxjs";
import {
  Page,
  PageTemplate,
  File,
  MarkdownSettings,
  CommandArguments,
} from "@models";
import {
  ArgumentsHelper,
  CliCommand,
  execScript,
  FileHelpers,
  FolderHelpers,
  ListHelpers,
  Logger,
  MarkdownHelper,
} from "@helpers";

const normalizeValue = (value?: string | null): string =>
  (value || "").toString().toLowerCase();

export class PagesHelper {
  private static pages: File[] = [];
  private static processedPages: { [slug: string]: number } = {};

  /**
   * Retrieve all the pages from the current site
   * @param webUrl
   */
  public static async getAllPages(webUrl: string): Promise<void> {
    PagesHelper.pages = await FileHelpers.getAllPages(webUrl, "sitepages");
    Logger.debug(`Existing pages`);
    Logger.debug(PagesHelper.pages);
  }

  /**
   * Cleaning up all the untouched pages
   * @param webUrl
   */
  public static async clean(
    webUrl: string,
    options: CommandArguments
  ): Promise<Observable<string>> {
    return new Observable((observer) => {
      (async () => {
        const untouched = this.getUntouchedPages().filter((slug) => {
          const normalizedSlug = normalizeValue(slug);
          return (
            !!normalizedSlug &&
            !normalizedSlug.startsWith("templates") &&
            normalizedSlug.endsWith(".aspx")
          );
        });
        Logger.debug(`Removing the following files`);
        Logger.debug(untouched);
        for (const slug of untouched) {
          try {
            if (slug) {
              Logger.debug(`Cleaning up page: ${slug}`);
              observer.next(`Cleaning up page: ${slug}`);
              const filePath = `sitepages/${slug}`;
              const relUrl = FileHelpers.getRelUrl(webUrl, filePath);
              await execScript<string>(
                ArgumentsHelper.parse(
                  `spo file remove --webUrl "${webUrl}" --url "${relUrl}" --force`
                ),
                CliCommand.getRetry()
              );
            }
          } catch (err) {
            const error = err instanceof Error ? err : new Error(err as any);
            observer.error(error);
            Logger.debug(error.message);

            if (!options.continueOnError) {
              throw error;
            }
          }
        }
        observer.complete();
      })();
    });
  }

  /**
   * Check if the page exists, and if it doesn't it will be created
   * @param webUrl
   * @param slug
   * @param title
   */
  public static async createPageIfNotExists(
    webUrl: string,
    slug: string,
    title: string,
    layout: string = "Article",
    commentsDisabled: boolean = false,
    description: string = "",
    template: string | null = null,
    skipExistingPages: boolean = false
  ): Promise<boolean> {
    try {
      const relativeUrl = FileHelpers.getRelUrl(webUrl, `sitepages/${slug}`);

      if (skipExistingPages) {
        if (PagesHelper.pages && PagesHelper.pages.length > 0) {
          const page = PagesHelper.pages.find((page: File) => {
            const fileRef = normalizeValue(page.FileRef);
            return (
              !!fileRef && fileRef === normalizeValue(relativeUrl)
            );
          });
          if (page) {
            // Page already existed
            PagesHelper.processedPages[normalizeValue(slug)] = page.ID;
            Logger.debug(
              `Processed pages: ${JSON.stringify(PagesHelper.processedPages)}`
            );
            return true;
          }
        }
      }

      let pageData: Page | string = await execScript(
        ArgumentsHelper.parse(
          `spo page get --webUrl "${webUrl}" --name "${slug}" --metadataOnly --output json`
        ),
        false
      );
      if (pageData && typeof pageData === "string") {
        pageData = JSON.parse(pageData);
      }

      PagesHelper.processedPages[normalizeValue(slug)] = (
        pageData as Page
      ).ListItemAllFields.Id;
      Logger.debug(
        `Processed pages: ${JSON.stringify(PagesHelper.processedPages)}`
      );

      Logger.debug(pageData);

      let cmdArgs = ``;

      if (pageData && (pageData as Page).title !== title) {
        cmdArgs = `--title "${title}"`;
      }

      if (pageData && description) {
        cmdArgs = `${cmdArgs} --description "${description}"`;
      }

      if (pageData && (pageData as Page).layoutType !== layout) {
        cmdArgs = `${cmdArgs} --layoutType "${layout}"`;
      }

      if (
        pageData &&
        (pageData as Page).commentsDisabled !== commentsDisabled
      ) {
        cmdArgs = `${cmdArgs} --commentsEnabled ${
          commentsDisabled ? "false" : "true"
        }`;
      }

      if (cmdArgs) {
        await execScript(
          ArgumentsHelper.parse(
            `spo page set --webUrl "${webUrl}" --name "${slug}" ${cmdArgs}`
          ),
          CliCommand.getRetry()
        );
      }

      return true;
    } catch (e) {
      // Check if folders for the file need to be created
      if (slug.split("/").length > 1) {
        const folders = slug.split("/");
        await FolderHelpers.create(
          "sitepages",
          folders.slice(0, folders.length - 1),
          webUrl
        );
      }

      if (template) {
        let templates: PageTemplate[] | string = await execScript(
          ArgumentsHelper.parse(
            `spo page template list --webUrl "${webUrl}" --output json`
          ),
          CliCommand.getRetry()
        );
        if (templates && typeof templates === "string") {
          templates = JSON.parse(templates);
        }

        Logger.debug(templates);

        const pageTemplate = (templates as PageTemplate[]).find(
          (t) => t.Title === template
        );
        if (pageTemplate) {
          const templateUrl = normalizeValue(pageTemplate.Url).replace(
            "sitepages/",
            ""
          );
          await execScript(
            ArgumentsHelper.parse(
              `spo page copy --webUrl "${webUrl}" --sourceName "${templateUrl}" --targetUrl "${slug}"`
            ),
            CliCommand.getRetry()
          );
          await execScript(
            ArgumentsHelper.parse(
              `spo page set --webUrl "${webUrl}" --name "${slug}" --publish`
            ),
            CliCommand.getRetry()
          );
          return await this.createPageIfNotExists(
            webUrl,
            slug,
            title,
            layout,
            commentsDisabled,
            description,
            null,
            skipExistingPages
          );
        } else {
          console.log(
            `Template "${template}" not found on the site, will create a default page instead.`
          );
        }
      }

      // File doesn't exist
      await execScript(
        ArgumentsHelper.parse(
          `spo page add --webUrl "${webUrl}" --name "${slug}" --title "${title}" --layoutType "${layout}" ${
            commentsDisabled ? "" : "--commentsEnabled"
          } --description "${description}"`
        ),
        CliCommand.getRetry()
      );

      return false;
    }
  }

  /**
   * Retrieve all the page controls
   * @param webUrl
   * @param slug
   */
  public static async getPageControls(
    webUrl: string,
    slug: string
  ): Promise<string> {
    Logger.debug(`Get page controls for ${slug}`);

    let output = await execScript<any | string>(
      ArgumentsHelper.parse(
        `spo page get --webUrl "${webUrl}" --name "${slug}" --output json`
      ),
      CliCommand.getRetry()
    );
    if (output && typeof output === "string") {
      output = JSON.parse(output);
    }

    Logger.debug(JSON.stringify(output.canvasContentJson || "[]"));
    return output.canvasContentJson || "[]";
  }

  /**
   * Ensure the page contains at least one default section
   * @param webUrl
   * @param slug
   */
  public static async ensureDefaultSection(
    webUrl: string,
    slug: string
  ): Promise<void> {
    Logger.debug(`Ensuring default section exists for ${slug}`);
    await execScript(
      ArgumentsHelper.parse(
        `spo page section add --webUrl "${webUrl}" --pageName "${slug}" --sectionTemplate OneColumn`
      ),
      CliCommand.getRetry()
    );
  }

  /**
   * Inserts or create the control
   * @param webPartTitle
   * @param markdown
   */
  public static async insertOrCreateControl(
    webPartTitle: string,
    markdown: string,
    slug: string,
    webUrl: string,
    options: CommandArguments,
    wpId: string = null,
    mdOptions: MarkdownSettings | null,
    wasAlreadyParsed: boolean = false
  ) {
    Logger.debug(
      `Insert the markdown webpart for the page ${slug} - Control ID: ${wpId} - Was already parsed: ${wasAlreadyParsed}`
    );

    const wpData = await MarkdownHelper.getJsonData(
      webPartTitle,
      markdown,
      mdOptions,
      options,
      wasAlreadyParsed
    );

    if (wpId) {
      // Web part needs to be updated
      await execScript(
        ArgumentsHelper.parse(
          `spo page control set --webUrl "${webUrl}" --pageName "${slug}" --id "${wpId}" --webPartData @${wpData}`
        ),
        CliCommand.getRetry()
      );
    } else {
      // Add new markdown web part
      await execScript(
        ArgumentsHelper.parse(
          `spo page clientsidewebpart add --webUrl "${webUrl}" --pageName "${slug}" --webPartId 1ef5ed11-ce7b-44be-bc5e-4abd55101d16 --webPartData @${wpData}`
        ),
        CliCommand.getRetry()
      );
    }
  }

  /**
   * Set the page its metadata
   * @param webUrl
   * @param slug
   * @param metadata
   */
  public static async setPageMetadata(
    webUrl: string,
    slug: string,
    metadata: { [fieldName: string]: any } = null
  ) {
    const pageId = await this.getPageId(webUrl, slug);
    const pageList = await ListHelpers.getSitePagesList(webUrl);
    if (pageId && pageList) {
      let metadataCommand: string = `spo listitem set --listId "${pageList.Id}" --id ${pageId} --webUrl "${webUrl}"`;

      if (metadata) {
        for (const fieldName in metadata) {
          metadataCommand = `${metadataCommand} --${fieldName} "${metadata[fieldName]}"`;
        }
      }

      await execScript(
        ArgumentsHelper.parse(metadataCommand),
        CliCommand.getRetry()
      );
    }
  }

  /**
   * Set the page its description
   * @param webUrl
   * @param slug
   * @param description
   */
  public static async setPageDescription(
    webUrl: string,
    slug: string,
    description: string
  ) {
    const pageId = await this.getPageId(webUrl, slug);
    const pageList = await ListHelpers.getSitePagesList(webUrl);
    if (pageId && pageList) {
      await execScript(
        ArgumentsHelper.parse(
          `spo listitem set --listId "${pageList.Id}" --id ${pageId} --webUrl "${webUrl}" --Description "${description}" --systemUpdate`
        ),
        CliCommand.getRetry()
      );
    }
  }

  /**
   * Publish the page
   * @param webUrl
   * @param slug
   */
  public static async publishPageIfNeeded(webUrl: string, slug: string) {
    const relativeUrl = FileHelpers.getRelUrl(webUrl, `sitepages/${slug}`);
    const requiresCheckIn = await this.pageRequiresCheckIn(
      webUrl,
      relativeUrl
    );
    if (requiresCheckIn) {
      try {
        await execScript(
          ArgumentsHelper.parse(
            `spo file checkin --webUrl "${webUrl}" --url "${relativeUrl}"`
          ),
          false
        );
      } catch (e) {
        Logger.debug(
          `Page check-in skipped for ${relativeUrl}: ${e instanceof Error ? e.message : e}`
        );
      }
    } else {
      Logger.debug(`Skipping check-in for ${relativeUrl}, no pending checkout.`);
    }
    await execScript(
      ArgumentsHelper.parse(
        `spo page set --name "${slug}" --webUrl "${webUrl}" --publish`
      ),
      CliCommand.getRetry()
    );
  }

  /**
   * Retrieve the page id
   * @param webUrl
   * @param slug
   */
  private static async getPageId(webUrl: string, slug: string) {
    const normalizedSlug = normalizeValue(slug);
    if (!PagesHelper.processedPages[normalizedSlug]) {
      let pageData: any = await execScript(
        ArgumentsHelper.parse(
          `spo page get --webUrl "${webUrl}" --name "${slug}" --metadataOnly --output json`
        ),
        CliCommand.getRetry()
      );
      if (pageData && typeof pageData === "string") {
        pageData = JSON.parse(pageData);

        Logger.debug(pageData);

        if (pageData.ListItemAllFields && pageData.ListItemAllFields.Id) {
          PagesHelper.processedPages[normalizedSlug] =
            pageData.ListItemAllFields.Id;
          return PagesHelper.processedPages[normalizedSlug];
        }

        return null;
      }
    }

    return PagesHelper.processedPages[normalizedSlug];
  }

  private static async pageRequiresCheckIn(
    webUrl: string,
    relativeUrl: string
  ): Promise<boolean> {
    try {
      let fileInfo: any = await execScript(
        ArgumentsHelper.parse(
          `spo file get --webUrl "${webUrl}" --url "${relativeUrl}" -o json`
        ),
        CliCommand.getRetry()
      );
      if (fileInfo && typeof fileInfo === "string") {
        fileInfo = JSON.parse(fileInfo);
      }
      if (!fileInfo) {
        return false;
      }

      const checkOutType = fileInfo.CheckOutType;
      const checkedOutBy =
        fileInfo.CheckedOutByUserId ||
        fileInfo.CheckedOutByUser ||
        fileInfo.LockedByUserId;

      if (typeof checkOutType !== "undefined" && checkOutType !== null) {
        if (
          (typeof checkOutType === "number" && checkOutType !== 2) ||
          (typeof checkOutType === "string" &&
            checkOutType.toLowerCase() !== "none")
        ) {
          return true;
        }
      }

      return !!checkedOutBy;
    } catch (err) {
      Logger.debug(
        `Unable to determine checkout status for ${relativeUrl}: ${
          err instanceof Error ? err.message : err
        }`
      );
    }

    return false;
  }

  /**
   * Receive all the pages which have not been touched
   */
  private static getUntouchedPages(): string[] {
    let untouched: string[] = [];
    for (const page of PagesHelper.pages) {
      const { FileRef: url } = page;
      const normalizedUrl = normalizeValue(url);
      const slug = normalizedUrl.split("/sitepages/")[1];
      const normalizedSlug = normalizeValue(slug);
      if (normalizedSlug && !PagesHelper.processedPages[normalizedSlug]) {
        untouched.push(normalizedSlug);
      }
    }
    return untouched;
  }
}
