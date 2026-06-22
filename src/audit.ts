import { chromium, Page } from "playwright";
import ExcelJS from "exceljs";
import * as path from "path";
import * as fs from "fs";

// Configuration
const BASE_URL = "http://pm.solusi-ku.id";
const INITIAL_URL = `${BASE_URL}/zentao/execution-task-34-all-0-status,id_desc-17-100.html`;
const EXECUTION_LIST_URL = `${BASE_URL}/zentao/project-execution-all-33-order_asc-0-2-100-1.html`;

const USERNAME = "dystian.en";
const PASSWORD = "Solusiku123";

interface TicketInfo {
  id: string;
  iteration: string;
  title: string;
  taskStatus: string;
  status: "Has Description" | "Missing Description";
  url: string;
}

// Helper to navigate with retry logic
async function ensureLoggedIn(page: Page) {
  const accountInput = page.locator("#account");
  if (await accountInput.isVisible()) {
    console.log(
      "[INFO] Login page detected (possibly expired session). Re-authenticating...",
    );

    // Switch language to English if toggle is visible
    const langToggle = page.locator("#langs-toggle");
    if (await langToggle.isVisible()) {
      await langToggle.click({ force: true });

      const englishOption = page.locator(
        'a[data-call="switchLang"][data-params="en"]',
      );
      await englishOption.waitFor({ state: "visible", timeout: 5000 });
      await englishOption.click({ force: true });

      // Wait until the lang toggle text changes to "English"
      await page
        .waitForFunction(
          () => {
            const toggle = document.querySelector("#langs-toggle");
            return toggle?.textContent?.trim().toLowerCase().includes("en");
          },
          { timeout: 10000 },
        )
        .catch(() =>
          console.warn("[WARN] Language toggle did not confirm switch to EN"),
        );

      // Additional buffer for any re-render triggered by lang change
      await page.waitForTimeout(500);
    }

    await page.fill("#account", USERNAME);
    await page.fill("#password", PASSWORD);
    await page.locator("#submit").click({ force: true });
    await page
      .waitForLoadState("networkidle", { timeout: 10000 })
      .catch(() => {});
    await page.waitForTimeout(5000);
    console.log("[INFO] Re-authentication completed.");
  }
}

async function navigateWithRetry(
  page: Page,
  url: string,
  retries = 3,
): Promise<void> {
  for (let i = 0; i < retries; i++) {
    try {
      console.log(
        `[INFO] Navigating to ${url} (Attempt ${i + 1}/${retries})...`,
      );
      await page.goto(url, { waitUntil: "load", timeout: 45000 });
      // Gracefully wait for network idle for up to 10 seconds, but do not fail if it times out
      await page
        .waitForLoadState("networkidle", { timeout: 10000 })
        .catch(() => {
          console.log("[INFO] Network idle timeout, proceeding.");
        });
      // Check if session expired and re-authenticate if necessary
      await ensureLoggedIn(page);
      return;
    } catch (error) {
      console.error(
        `[WARN] Navigation to ${url} failed:`,
        error instanceof Error ? error.message : error,
      );
      if (i === retries - 1) {
        throw new Error(`Failed to load ${url} after ${retries} attempts.`);
      }
      await page.waitForTimeout(5000); // Wait 5s before retry
    }
  }
}

// Clean HTML and validate description
function validateDescription(html: string): boolean {
  if (!html) return false;
  // Remove HTML tags
  const textWithoutTags = html.replace(/<[^>]*>/g, " ");
  // Decode HTML entities
  const decoded = textWithoutTags
    .replace(/&nbsp;/g, " ")
    .replace(/&#160;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#\d+;/g, " ")
    .replace(/&[a-z0-9#]+;/gi, " ");
  // Clean whitespace
  const cleaned = decoded.replace(/\s+/g, " ").trim();
  const lowerCleaned = cleaned
    .toLowerCase()
    .replace(/[.\s]+/g, " ")
    .trim();
  if (
    lowerCleaned === "no describe" ||
    lowerCleaned === "no description" ||
    lowerCleaned === "暂无描述" ||
    lowerCleaned.includes("暂无描述")
  ) {
    return false;
  }
  return cleaned.length > 0;
}

function mapStatus(status: string): string {
  if (!status) return "";
  const s = status.toLowerCase();
  switch (s) {
    case "wait":
      return "Wait";
    case "doing":
      return "On Progress";
    case "done":
      return "Resolved";
    case "closed":
      return "Closed";
    case "cancel":
      return "Canceled";
    default:
      return status.charAt(0).toUpperCase() + status.slice(1);
  }
}

async function runAudit() {
  console.log("[INFO] Starting Zentao Ticket Description Audit Bot...");

  const browser = await chromium.launch({
    headless: process.env.HEADLESS === "false",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
  });
  const page = await context.newPage();

  try {
    // 1. Open Zentao Initial Page
    await navigateWithRetry(page, INITIAL_URL);

    // Check if we are still on the login page (if navigateWithRetry did not already log us in)
    const accountInput = page.locator("#account");
    if (await accountInput.isVisible()) {
      // 2. Change Language to English
      console.log("[INFO] Attempting to change language to English...");
      const langToggle = page.locator("#langs-toggle");
      if (await langToggle.isVisible()) {
        await langToggle.click({ force: true });

        const englishOption = page.locator(
          'a[data-call="switchLang"][data-params="en"]',
        );
        if (await englishOption.isVisible()) {
          await englishOption.click({ force: true });
          console.log("[INFO] Language switched to English.");
          await page
            .waitForLoadState("networkidle", { timeout: 10000 })
            .catch(() => {});
        } else {
          console.log(
            "[INFO] English option not visible, assuming already English or switched.",
          );
        }
      } else {
        console.log(
          "[INFO] Language toggle not visible, assuming already in English.",
        );
      }

      // 3. Login
      console.log("[INFO] Logging in...");
      await page.fill("#account", USERNAME);
      await page.fill("#password", PASSWORD);
      await page.locator("#submit").click({ force: true });
      await page
        .waitForLoadState("networkidle", { timeout: 10000 })
        .catch(() => {});
      await page.waitForTimeout(5000); // Allow post-login redirections to fully settle
      console.log("[INFO] Login completed.");
    } else {
      console.log(
        "[INFO] Already logged in (re-authentication handled by navigation).",
      );
    }

    // 4. Open Execution List
    console.log("[INFO] Navigating to Execution List...");
    await navigateWithRetry(page, EXECUTION_LIST_URL);

    const frame = page.frameLocator('iframe[id^="appIframe-"]').first();

    // Enable Show Task if not checked inside the iframe
    const showTaskCheckbox = frame.locator("#showTask");
    if ((await showTaskCheckbox.count()) > 0) {
      const isChecked = await showTaskCheckbox.isChecked();
      if (!isChecked) {
        console.log('[INFO] Checking "Show Task" checkbox...');
        await frame.locator('label[for="showTask"]').click();
        await page
          .waitForLoadState("networkidle", { timeout: 10000 })
          .catch(() => {});
        await page.waitForTimeout(3000);
      } else {
        console.log('[INFO] "Show Task" checkbox is already checked.');
      }
    } else {
      console.log(
        '[INFO] "Show Task" checkbox locator check skipped (not visible or already loaded).',
      );
    }

    // 5. Find Iterations (January 2026 - June 2026)
    console.log(
      "[INFO] Scanning for iterations between January and June 2026...",
    );
    console.log("[INFO] Waiting for iterations table to load...");
    const dtableSelector = "div.dtable, #table-project-execution";
    await frame
      .locator(dtableSelector)
      .first()
      .waitFor({ state: "visible", timeout: 30000 })
      .catch(() => {
        console.log("[WARN] Timeout waiting for iterations table.");
      });
    await page.waitForTimeout(3000); // Give it a bit of time to settle

    const iterationsToProcess = await frame.locator("body").evaluate(() => {
      const elm = document.querySelector(
        "div.dtable, #table-project-execution",
      );
      if (!elm) return [];
      const dtableInstance = (window as any).zui?.DTable?.query(elm);
      if (!dtableInstance) return [];
      const rows = dtableInstance.options?.data || dtableInstance.data || [];

      const monthsRegex = /(01|02|03|04|05|06)\/2026/;
      const results: { name: string; url: string }[] = [];

      rows.forEach((r: any) => {
        const name = r.name || r.data?.name || "";
        const rawID = r.rawID || r.data?.rawID || r.id || "";
        if (monthsRegex.test(name) && rawID) {
          results.push({
            name: name.trim(),
            url: `/zentao/execution-task-${rawID}.html`,
          });
        }
      });
      return results;
    });

    // Limit iterations to process for testing
    const LIMIT_ITERATIONS = 2; // Set to null to process all iterations
    const iterationsToRun = LIMIT_ITERATIONS
      ? iterationsToProcess.slice(0, LIMIT_ITERATIONS)
      : iterationsToProcess;

    console.log(
      `[INFO] Found ${iterationsToProcess.length} iterations. Testing with first ${iterationsToRun.length} iterations...`,
    );
    for (const iter of iterationsToRun) {
      // Convert to absolute URL relative to current page URL
      iter.url = iter.url.startsWith("http")
        ? iter.url
        : new URL(iter.url, page.url()).toString();
      console.log(`  - ${iter.name} (${iter.url})`);
    }

    // 6. Crawl Tickets inside Iterations
    const processedTickets = new Set<string>();
    const allTickets: TicketInfo[] = [];

    for (const iteration of iterationsToRun) {
      console.log(`\n[INFO] Iteration : ${iteration.name}`);

      // Go to iteration page
      await navigateWithRetry(page, iteration.url);

      // Wait for tasks table in the iframe
      const taskDTableSelector = "div.dtable, #table-execution-task";
      await frame
        .locator(taskDTableSelector)
        .first()
        .waitFor({ state: "visible", timeout: 30000 })
        .catch(() => {
          console.log("[WARN] Timeout waiting for task table.");
        });
      await page.waitForTimeout(3000);

      // Change filter to "All"
      const allFilter = frame.locator('a[data-id="all"]');
      if ((await allFilter.count()) > 0) {
        console.log("[INFO] Clicking 'All' filter tab...");
        await allFilter.first().click({ force: true });
        await page.waitForTimeout(3000);
        // Wait for tasks table to reload
        await frame
          .locator(taskDTableSelector)
          .first()
          .waitFor({ state: "visible", timeout: 30000 })
          .catch(() => {
            console.log(
              "[WARN] Timeout waiting for task table after clicking 'All'.",
            );
          });
        await page.waitForTimeout(2000);
      } else {
        console.log("[WARN] 'All' filter tab not found.");
      }

      // Extract tickets via JS
      const iterationTicketsToVisit = await frame
        .locator("body")
        .evaluate(() => {
          const elm = document.querySelector(
            "div.dtable, #table-execution-task",
          );
          if (!elm) return [];
          const dtableInstance = (window as any).zui?.DTable?.query(elm);
          if (!dtableInstance) return [];
          const rows =
            dtableInstance.options?.data || dtableInstance.data || [];

          return rows
            .map((r: any) => {
              const id = String(r.rawID || r.id || "");
              const title = String(r.name || r.title || "");
              const status = String(
                r.status || r.data?.status || r.taskStatus || "",
              );
              return { id, title, status };
            })
            .filter((t: any) => t.id);
        });

      console.log(
        `[INFO] Found ${iterationTicketsToVisit.length} tickets in iteration: ${iteration.name}`,
      );

      // Visit each ticket
      for (const ticket of iterationTicketsToVisit) {
        if (processedTickets.has(ticket.id)) {
          continue;
        }
        processedTickets.add(ticket.id);

        const ticketUrl = `${BASE_URL}/zentao/task-view-${ticket.id}.html`;

        try {
          await navigateWithRetry(page, ticketUrl);

          // Check for Task Description (in iframe or top-level)
          let articleHtml = "";
          const iframeSelector = 'iframe[id^="appIframe-"]';
          const hasIframe = (await page.locator(iframeSelector).count()) > 0;

          if (hasIframe) {
            const activeIframe = page.frameLocator(iframeSelector).first();
            const descSectionIframe = activeIframe.locator(
              'div[zui-key="Task Description"], .detail-section[zui-key="Task Description"], .detail-section',
            );

            // Wait for description section to load inside iframe
            await descSectionIframe
              .first()
              .waitFor({ state: "visible", timeout: 15000 })
              .catch(() => {
                console.log(
                  "[WARN] Timeout waiting for description section inside iframe.",
                );
              });

            if ((await descSectionIframe.count()) > 0) {
              const article = activeIframe.locator(".article").first();
              if ((await article.count()) > 0) {
                articleHtml = await article.innerHTML();
              }
            } else {
              const fallbackArticleIframe = activeIframe
                .locator(".detail-section-content .article")
                .first();
              if ((await fallbackArticleIframe.count()) > 0) {
                articleHtml = await fallbackArticleIframe.innerHTML();
              }
            }
          } else {
            const descSectionTop = page.locator(
              'div[zui-key="Task Description"], .detail-section[zui-key="Task Description"], .detail-section',
            );

            // Wait for description section to load on top-level
            await descSectionTop
              .first()
              .waitFor({ state: "visible", timeout: 15000 })
              .catch(() => {
                console.log(
                  "[WARN] Timeout waiting for description section on top-level.",
                );
              });

            if ((await descSectionTop.count()) > 0) {
              const article = descSectionTop.locator(".article").first();
              if ((await article.count()) > 0) {
                articleHtml = await article.innerHTML();
              }
            } else {
              const fallbackArticleTop = page
                .locator(".detail-section-content .article")
                .first();
              if ((await fallbackArticleTop.count()) > 0) {
                articleHtml = await fallbackArticleTop.innerHTML();
              }
            }
          }

          const hasDescription = validateDescription(articleHtml);
          const status = hasDescription
            ? "Has Description"
            : "Missing Description";

          console.log(`[INFO] Ticket    : ${ticket.id}`);
          console.log(`[INFO] Status    : ${status}`);

          allTickets.push({
            id: ticket.id,
            iteration: iteration.name,
            title: ticket.title,
            taskStatus: mapStatus(ticket.status),
            status,
            url: ticketUrl,
          });

          // Polite crawling delay
          await page.waitForTimeout(500);
        } catch (ticketError) {
          console.error(
            `[ERROR] Failed to process ticket ${ticket.id}:`,
            ticketError instanceof Error ? ticketError.message : ticketError,
          );
          allTickets.push({
            id: ticket.id,
            iteration: iteration.name,
            title: ticket.title,
            taskStatus: mapStatus(ticket.status),
            status: "Missing Description", // treat as missing/unreachable
            url: ticketUrl,
          });
        }
      }
    }

    // 7. Generate Excel Report
    const totalIterationsCount = iterationsToProcess.length;
    const totalTicketsCount = allTickets.length;
    const hasDescCount = allTickets.filter(
      (t) => t.status === "Has Description",
    ).length;
    const missingDescCount = allTickets.filter(
      (t) => t.status === "Missing Description",
    ).length;

    const completionRate =
      totalTicketsCount > 0 ? (hasDescCount / totalTicketsCount) * 100 : 0;
    const missingRate =
      totalTicketsCount > 0 ? (missingDescCount / totalTicketsCount) * 100 : 0;

    // Display summary in console
    console.log("\n[SUMMARY]");
    console.log(`Total Iteration : ${totalIterationsCount}`);
    console.log(`Total Ticket : ${totalTicketsCount}`);
    console.log(`Has Description : ${hasDescCount}`);
    console.log(`Missing Description : ${missingDescCount}`);

    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Summary
    const summarySheet = workbook.addWorksheet("Summary");
    summarySheet.getColumn("A").width = 30;
    summarySheet.getColumn("B").width = 20;

    // Add visual design to Summary sheet
    summarySheet.addRow(["Metric", "Value"]);
    summarySheet.addRow(["Total Iteration", totalIterationsCount]);
    summarySheet.addRow(["Total Ticket", totalTicketsCount]);
    summarySheet.addRow(["Ticket Dengan Description", hasDescCount]);
    summarySheet.addRow(["Ticket Tanpa Description", missingDescCount]);
    summarySheet.addRow(["Completion Rate", `${completionRate.toFixed(2)}%`]);
    summarySheet.addRow(["Missing Rate", `${missingRate.toFixed(2)}%`]);

    // Apply header style to Summary sheet
    const summaryHeaderRow = summarySheet.getRow(1);
    summaryHeaderRow.font = {
      name: "Arial",
      family: 2,
      size: 11,
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
    summaryHeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1F497D" }, // Navy blue
    };
    summaryHeaderRow.alignment = { vertical: "middle", horizontal: "center" };

    // Apply borders and styling to summary rows
    for (let r = 2; r <= 7; r++) {
      const row = summarySheet.getRow(r);
      row.getCell("A").font = { bold: true };
      row.getCell("A").alignment = { horizontal: "left" };
      row.getCell("B").alignment = { horizontal: "right" };

      // Color coding for rate values
      if (r === 6) {
        row.getCell("B").font = { color: { argb: "FF008000" }, bold: true }; // Green for completion rate
      } else if (r === 7) {
        row.getCell("B").font = { color: { argb: "FFFF0000" }, bold: true }; // Red for missing rate
      }

      // Add borders
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFD3D3D3" } },
          left: { style: "thin", color: { argb: "FFD3D3D3" } },
          bottom: { style: "thin", color: { argb: "FFD3D3D3" } },
          right: { style: "thin", color: { argb: "FFD3D3D3" } },
        };
      });
    }

    // Sheet 2: All Tickets
    const allTicketsSheet = workbook.addWorksheet("All Tickets");
    allTicketsSheet.columns = [
      { header: "Ticket ID", key: "id", width: 15 },
      { header: "Iteration", key: "iteration", width: 30 },
      { header: "Title", key: "title", width: 50 },
      { header: "Task Status", key: "taskStatus", width: 15 },
      { header: "Description Status", key: "status", width: 25 },
      { header: "URL", key: "url", width: 70 },
    ];

    // Add data to Sheet 2
    for (const ticket of allTickets) {
      allTicketsSheet.addRow({
        id: ticket.id,
        iteration: ticket.iteration,
        title: ticket.title,
        taskStatus: ticket.taskStatus,
        status: ticket.status,
        url: ticket.url,
      });
    }

    // Apply styles to Sheet 2 headers
    const allTicketsHeaderRow = allTicketsSheet.getRow(1);
    allTicketsHeaderRow.font = {
      name: "Arial",
      family: 2,
      size: 11,
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
    allTicketsHeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1F497D" },
    };
    allTicketsHeaderRow.alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    // Apply alternate row shading and status coloring to Sheet 2
    allTicketsSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      // Status text coloring (Description Status is now column E / 5)
      const statusCell = row.getCell("E");
      if (statusCell.value === "Has Description") {
        statusCell.font = { color: { argb: "FF008000" }, bold: true };
      } else {
        statusCell.font = { color: { argb: "FFFF0000" }, bold: true };
      }

      // Center align the Task Status column
      row.getCell("D").alignment = { horizontal: "center" };

      // Zebra striping
      if (rowNumber % 2 === 0) {
        row.eachCell((cell, colNumber) => {
          if (colNumber !== 5) {
            // Don't override status font colors
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFF2F5F8" },
            };
          }
        });
      }

      // Add borders
      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE6E6E6" } },
          left: { style: "thin", color: { argb: "FFE6E6E6" } },
          bottom: { style: "thin", color: { argb: "FFE6E6E6" } },
          right: { style: "thin", color: { argb: "FFE6E6E6" } },
        };
      });
    });

    // Sheet 3: Missing Description
    const missingDescSheet = workbook.addWorksheet("Missing Description");
    missingDescSheet.columns = [
      { header: "Ticket ID", key: "id", width: 15 },
      { header: "Iteration", key: "iteration", width: 30 },
      { header: "Title", key: "title", width: 50 },
      { header: "Task Status", key: "taskStatus", width: 15 },
      { header: "URL", key: "url", width: 70 },
    ];

    // Add data to Sheet 3
    const missingTickets = allTickets.filter(
      (t) => t.status === "Missing Description",
    );
    for (const ticket of missingTickets) {
      missingDescSheet.addRow({
        id: ticket.id,
        iteration: ticket.iteration,
        title: ticket.title,
        taskStatus: ticket.taskStatus,
        url: ticket.url,
      });
    }

    // Apply styles to Sheet 3 headers
    const missingHeaderRow = missingDescSheet.getRow(1);
    missingHeaderRow.font = {
      name: "Arial",
      family: 2,
      size: 11,
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
    missingHeaderRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFC00000" }, // Red header for missing description sheet
    };
    missingHeaderRow.alignment = { vertical: "middle", horizontal: "center" };

    // Apply alternate row shading and borders to Sheet 3
    missingDescSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      // Center align the Task Status column
      row.getCell("D").alignment = { horizontal: "center" };

      if (rowNumber % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFF0F0" }, // Light reddish tint for alternate rows
          };
        });
      }

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE6E6E6" } },
          left: { style: "thin", color: { argb: "FFE6E6E6" } },
          bottom: { style: "thin", color: { argb: "FFE6E6E6" } },
          right: { style: "thin", color: { argb: "FFE6E6E6" } },
        };
      });
    });

    // Write file
    const date = new Date();
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    let reportFilename = `zentao_description_audit_${yyyy}${mm}${dd}.xlsx`;
    let outputPath = path.join(process.cwd(), reportFilename);

    try {
      await workbook.xlsx.writeFile(outputPath);
      console.log(
        `\n[INFO] Excel report successfully generated at: ${outputPath}`,
      );
    } catch (writeError: any) {
      if (writeError.code === "EBUSY") {
        const hhmmss =
          String(date.getHours()).padStart(2, "0") +
          String(date.getMinutes()).padStart(2, "0") +
          String(date.getSeconds()).padStart(2, "0");
        reportFilename = `zentao_description_audit_${yyyy}${mm}${dd}_${hhmmss}.xlsx`;
        outputPath = path.join(process.cwd(), reportFilename);
        console.log(
          `[WARN] The file ${reportFilename} was busy or locked (EBUSY). Saving as backup: ${reportFilename}`,
        );
        await workbook.xlsx.writeFile(outputPath);
        console.log(
          `\n[INFO] Excel report successfully generated at: ${outputPath}`,
        );
      } else {
        throw writeError;
      }
    }
  } catch (error) {
    console.error("[FATAL ERROR] Audit process failed:", error);
    try {
      const errorScreenshotPath = path.join(
        process.cwd(),
        "error_screenshot.png",
      );
      await page.screenshot({ path: errorScreenshotPath, fullPage: true });
      console.log(`[INFO] Saved error screenshot to: ${errorScreenshotPath}`);
    } catch (e) {
      console.error("[WARN] Failed to capture error screenshot:", e);
    }
  } finally {
    await browser.close();
    console.log("[INFO] Browser closed. Process finished.");
  }
}

// Execute the audit
runAudit();
