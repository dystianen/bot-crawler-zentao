import { chromium, Page, Frame } from "playwright";
import ExcelJS from "exceljs";
import * as path from "path";
import * as fs from "fs";

// Simple .env loader
function loadEnv() {
  const envPath = path.join(process.cwd(), ".env");
  if (fs.existsSync(envPath)) {
    const content = fs.readFileSync(envPath, "utf-8");
    for (const line of content.split("\n")) {
      const match = line.match(/^\s*([\w.-]+)\s*=\s*(.*)\s*$/);
      if (match) {
        const key = match[1];
        let val = match[2].trim();
        // Remove quotes if present
        if (
          (val.startsWith('"') && val.endsWith('"')) ||
          (val.startsWith("'") && val.endsWith("'"))
        ) {
          val = val.slice(1, -1);
        }
        if (!process.env[key]) {
          process.env[key] = val;
        }
      }
    }
  }
}
loadEnv();

// Configuration
const BASE_URL = "http://pm.solusi-ku.id";
const INITIAL_URL = `${BASE_URL}/zentao/execution-task-34-all-0-status,id_desc-17-100.html`;
const EXECUTION_LIST_URL = `${BASE_URL}/zentao/project-execution-all-33-order_asc-0-2-100-1.html`;

const USERNAME = "ricky.widianto";
const PASSWORD = "Solusiku123";

interface TicketInfo {
  id: string;
  iteration: string;
  title: string;
  taskStatus: string;
  status:
    | "Has Description"
    | "Missing Description"
    | "Has Description (Auto Edited)";
  url: string;
}

// ─── Login helper ──────────────────────────────────────────────────────────────
async function login(page: Page) {
  console.log("[INFO] Logging in...");

  const langToggle = page.locator("#langs-toggle");
  if (await langToggle.isVisible({ timeout: 3000 }).catch(() => false)) {
    await langToggle.click({ force: true });
    const englishOption = page.locator(
      'a[data-call="switchLang"][data-params="en"]',
    );
    await englishOption.waitFor({ state: "visible", timeout: 5000 });
    await englishOption.click({ force: true });
    await page
      .waitForFunction(
        () =>
          document
            .querySelector("#langs-toggle")
            ?.textContent?.trim()
            .toLowerCase()
            .includes("en"),
        { timeout: 8000 },
      )
      .catch(() => {});
    await page.waitForTimeout(300);
  }

  await page.fill("#account", USERNAME);
  await page.fill("#password", PASSWORD);
  await page.locator("#submit").click({ force: true });

  await page
    .waitForFunction(() => !document.querySelector("#account"), {
      timeout: 20000,
    })
    .catch(() => {});

  console.log("[INFO] Login completed.");
}

// ─── Navigate helper ───────────────────────────────────────────────────────────
async function navigateTo(page: Page, url: string): Promise<void> {
  console.log(`[INFO] Navigating to ${url}...`);
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });

  // Check login — reduced timeout from 2000ms to 500ms
  const isLoginPage = await page
    .locator("#account")
    .isVisible({ timeout: 500 })
    .catch(() => false);
  if (isLoginPage) {
    console.log("[INFO] Session expired — re-authenticating...");
    await login(page);
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  }
}

// ─── Frame helper for execution list ───────────────────────────────────────────
function getAppIframe(page: Page) {
  return page.frameLocator('iframe[id^="appIframe-"]').first();
}

// ─── FAST description extractor ────────────────────────────────────────────────
// Uses page.$() + contentFrame() + single evaluate() instead of
// frameLocator + locator chains + multiple .count() calls
async function extractDescriptionFast(page: Page): Promise<string> {
  // 1. Wait for iframe element OR detail-section to appear in DOM
  await page
    .waitForSelector('iframe[id^="appIframe-"], .detail-section', {
      state: "attached",
      timeout: 5000,
    })
    .catch(() => {});

  // 2. Try to get the iframe's content frame
  const iframeHandle = await page.$('iframe[id^="appIframe-"]');
  let frame: Frame | null = null;

  if (iframeHandle) {
    // Retry contentFrame() — iframe might not have finished loading
    for (let i = 0; i < 10 && !frame; i++) {
      frame = await iframeHandle.contentFrame();
      if (!frame) await page.waitForTimeout(150);
    }
  }

  // 3. Pick the target: frame if available, otherwise main page
  const target: Page | Frame = frame || page;

  // 4. Wait for detail-sections to render inside target
  await target
    .waitForSelector(".detail-section", { timeout: 3000 })
    .catch(() => {});

  // 5. Single evaluate() call to extract description HTML
  return target.evaluate(() => {
    for (const section of document.querySelectorAll(".detail-section")) {
      const key = section.getAttribute("zui-key");
      const titleEl = section.querySelector(".detail-section-title");
      const isTaskDesc =
        key === "Task Description" ||
        titleEl?.textContent?.trim().toLowerCase() === "task description";
      if (isTaskDesc) {
        const article = section.querySelector(".article");
        return article ? article.innerHTML : "";
      }
    }
    return "";
  });
}

// ─── Description validator ─────────────────────────────────────────────────────
function validateDescription(html: string): boolean {
  if (!html) return false;
  const textWithoutTags = html.replace(/<[^>]*>/g, " ");
  const decoded = textWithoutTags
    .replace(/&nbsp;/g, " ")
    .replace(/&#160;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#\d+;/g, " ")
    .replace(/&[a-z0-9#]+;/gi, " ");
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
  )
    return false;
  return cleaned.length > 0;
}

function mapStatus(status: string): string {
  if (!status) return "";
  switch (status.toLowerCase()) {
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

// ─── Gemini Description Generator ──────────────────────────────────────────────
async function generateDescriptionWithGemini(
  title: string,
): Promise<string | null> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log(
      "[WARN] GEMINI_API_KEY is not defined in environment or .env file.",
    );
    return null;
  }

  console.log(
    `[INFO] Generating description using Gemini for task: "${title}"...`,
  );
  try {
    const prompt = `You are a technical writer and project manager.
      Your task is to generate a task description for a ZenTao ticket.
      The ticket title is: "${title}".

      The generated description MUST have exactly these 3 sections:
      1. Background
      2. Objective
      3. Expectation

      IMPORTANT FORMATTING RULES:
      - Write the description in ENGLISH language.
      - Use clean HTML formatting.
      - Section titles (Background, Objective, Expectation) must NOT have numbering. Just bold text.
      - If a section has MULTIPLE points, each point MUST be on its OWN line as a numbered list inside an <ol><li> tag.
      - If a section has only ONE point, write it as a normal <p> paragraph (no list needed).
      - Add an EMPTY line (empty <p></p> paragraph) BETWEEN each section for spacing.
      - Do NOT include any conversational text, markdown fence code block wraps (like \`\`\`html), or intro/outro.
      - Just return the raw HTML content.

      EXACT structure to follow (copy this format precisely):

      <p><strong>Background</strong></p>
      <p>[Single point as paragraph, OR multiple points as ordered list]</p>
      <p></p>
      <p><strong>Objective</strong></p>
      <p>[Single point as paragraph, OR multiple points as ordered list]</p>
      <p></p>
      <p><strong>Expectation</strong></p>
      <p>[Single point as paragraph, OR multiple points as ordered list]</p>

      Example with multiple points in each section:

      <p><strong>Background</strong></p>
      <p>The current system has several issues that need to be addressed:</p>
      <ol>
      <li>Users are experiencing timeout errors during peak hours.</li>
      <li>The database query for product listing is unoptimized.</li>
      <li>Error logging is incomplete, making debugging difficult.</li>
      </ol>
      <p></p>
      <p><strong>Objective</strong></p>
      <ol>
      <li>Optimize database queries to reduce response time below 200ms.</li>
      <li>Implement proper error handling and logging across all API endpoints.</li>
      </ol>
      <p></p>
      <p><strong>Expectation</strong></p>
      <ol>
      <li>All API endpoints respond within 200ms under normal load.</li>
      <li>Error logs capture full stack trace and request context.</li>
      <li>Unit tests and integration tests pass on staging environment.</li>
      </ol>`;

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          contents: [
            {
              parts: [
                {
                  text: prompt,
                },
              ],
            },
          ],
        }),
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Gemini API returned status ${response.status}: ${errorText}`,
      );
    }

    const data = (await response.json()) as any;
    let text = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (text) {
      // Clean up markdown block wraps if the AI accidentally returned them
      text = text
        .replace(/```html\s*/gi, "")
        .replace(/```\s*$/g, "")
        .trim();
      return text;
    }
    return null;
  } catch (error) {
    console.error("[ERROR] Failed to generate description with Gemini:", error);
    return null;
  }
}

// ─── Auto Edit Task Description ────────────────────────────────────────────────
async function autoEditTaskDescription(
  page: Page,
  ticketId: string,
  descriptionHtml: string,
): Promise<boolean> {
  // ── 1. Langsung navigate ke URL edit ──────────────────────────────────
  const editUrl = `${BASE_URL}/zentao/task-edit-${ticketId}.html`;
  console.log(`[INFO] Navigating to edit page: ${editUrl}...`);
  await navigateTo(page, editUrl);

  // ── 2. Resolve iframe (edit page juga di dalam iframe) ────────────────
  await page
    .waitForSelector('iframe[id^="appIframe-"], .ProseMirror', {
      state: "attached",
      timeout: 8000,
    })
    .catch(() => {});

  const iframeHandle = await page.$('iframe[id^="appIframe-"]');
  let frame: Frame | null = null;
  if (iframeHandle) {
    for (let i = 0; i < 15 && !frame; i++) {
      frame = await iframeHandle.contentFrame();
      if (!frame) await page.waitForTimeout(200);
    }
  }
  const target: Page | Frame = frame || page;

  // ── 3. Tunggu ProseMirror editor ──────────────────────────────────────
  const editorSelector = ".ProseMirror";
  await target.waitForSelector(editorSelector, { timeout: 10000 });

  console.log(`[INFO] Filling description editor...`);
  await target.locator(editorSelector).click(); // focus via click

  // ── 4. Inject HTML via clipboard paste ────────────────────────────────
  // innerHTML TIDAK bekerja karena ProseMirror punya internal doc state.
  // ClipboardEvent "paste" memicu transaction di ProseMirror.
  await target.locator(editorSelector).evaluate(async (el, html) => {
    const dt = new DataTransfer();
    dt.setData("text/html", html);
    dt.setData("text/plain", html.replace(/<[^>]*>/g, ""));

    const pasteEvent = new ClipboardEvent("paste", {
      bubbles: true,
      cancelable: true,
      clipboardData: dt,
    });
    el.dispatchEvent(pasteEvent);
  }, descriptionHtml);

  await page.waitForTimeout(500);

  // Verifikasi isi editor tidak kosong setelah paste
  const editorContent = await target
    .locator(editorSelector)
    .evaluate((el) => el.innerHTML);
  if (
    !editorContent ||
    editorContent.trim() === "<p><br></p>" ||
    editorContent.trim() === ""
  ) {
    console.log(`[WARN] Paste failed, trying execCommand fallback...`);
    await target.locator(editorSelector).click();
    await target.locator(editorSelector).evaluate((el, html) => {
      el.focus();
      document.execCommand("selectAll", false);
      document.execCommand("insertHTML", false, html);
    }, descriptionHtml);
    await page.waitForTimeout(300);
  }

  // ── 5. Klik tombol Save (type="submit") ──────────────────────────────
  // Selector yang benar: button[type="submit"] yang punya text "Save"
  // Ini menghindari tombol hidden team-saveBtn yang type="button"
  const submitBtn = target.locator('button[type="submit"]');

  // Filter yang visible dan mengandung text "Save"
  const visibleSaveBtn = submitBtn.filter({ hasText: "Save" });

  await visibleSaveBtn.waitFor({ state: "visible", timeout: 8000 });
  console.log(`[INFO] Clicking Save button...`);
  await visibleSaveBtn.click({ force: true });

  // ── 6. Tunggu save selesai ────────────────────────────────────────────
  console.log(`[INFO] Waiting for save to complete...`);
  await page.waitForTimeout(3000);

  // ── 7. Verifikasi di halaman view ─────────────────────────────────────
  const viewUrl = `${BASE_URL}/zentao/task-view-${ticketId}.html`;
  await navigateTo(page, viewUrl);
  const updatedHtml = await extractDescriptionFast(page);
  const isVerified = validateDescription(updatedHtml);

  if (isVerified) {
    console.log(
      `[SUCCESS] Description auto-edited and verified for ticket ${ticketId}.`,
    );
    return true;
  } else {
    console.log(`[WARN] Auto-edit verification failed for ticket ${ticketId}.`);
    return false;
  }
}

// ─── Main ──────────────────────────────────────────────────────────────────────
async function runAudit() {
  console.log("[INFO] Starting Zentao Ticket Description Audit Bot...");

  const browser = await chromium.launch({
    headless: false,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
  });
  const page = await context.newPage();

  try {
    // ── 1. Open initial URL (handles login if needed) ──────────────────────
    await navigateTo(page, INITIAL_URL);

    // ── 2. Open Execution List ─────────────────────────────────────────────
    console.log("[INFO] Navigating to Execution List...");
    await navigateTo(page, EXECUTION_LIST_URL);

    const frame = getAppIframe(page);

    const dtableSelector = "div.dtable, #table-project-execution";
    await frame
      .locator(dtableSelector)
      .first()
      .waitFor({ state: "visible", timeout: 30000 });

    // Enable Show Task if unchecked
    // const showTaskCheckbox = frame.locator("#showTask");
    // if ((await showTaskCheckbox.count()) > 0 && !(await showTaskCheckbox.isChecked())) {
    //   console.log('[INFO] Checking "Show Task" checkbox...');
    //   await frame.locator('label[for="showTask"]').click();
    //   await frame.locator(dtableSelector).first().waitFor({ state: "visible", timeout: 15000 });
    // }

    // ── 3. Find iterations Jan–Jun 2026 ────────────────────────────────────
    console.log(
      "[INFO] Scanning for iterations between January and June 2026...",
    );

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

    const LIMIT_ITERATIONS = null;
    const iterationsToRun = LIMIT_ITERATIONS
      ? iterationsToProcess.slice(0, LIMIT_ITERATIONS)
      : iterationsToProcess;

    console.log(
      `[INFO] Found ${iterationsToProcess.length} iterations. Running first ${iterationsToRun.length}...`,
    );
    for (const iter of iterationsToRun) {
      iter.url = iter.url.startsWith("http")
        ? iter.url
        : new URL(iter.url, page.url()).toString();
      console.log(`  - ${iter.name} (${iter.url})`);
    }

    // ── 4. Crawl tickets ───────────────────────────────────────────────────
    const processedTickets = new Set<string>();
    const allTickets: TicketInfo[] = [];
    const taskDTableSelector = "div.dtable, #table-execution-task";

    for (const iteration of iterationsToRun) {
      console.log(`\n[INFO] Iteration : ${iteration.name}`);

      await navigateTo(page, iteration.url);

      // Wait for task table using the app iframe (execution list context)
      const iterFrame = getAppIframe(page);
      await iterFrame
        .locator(taskDTableSelector)
        .first()
        .waitFor({ state: "visible", timeout: 30000 });

      // Click "All" filter
      await frame
        .locator(taskDTableSelector)
        .first()
        .waitFor({ state: "visible", timeout: 30000 });

      // Click "All" filter
      const allFilter = frame.locator('a[data-id="all"]');
      if ((await allFilter.count()) > 0) {
        console.log("[INFO] Clicking 'All' filter tab...");
        await allFilter.first().click({ force: true });
        // Wait for table to reflect the filter change
        await frame
          .locator(taskDTableSelector)
          .first()
          .waitFor({ state: "visible", timeout: 15000 });
        await page.waitForTimeout(500); // tiny settle for JS re-render
      }

      const iterationTickets = await iterFrame.locator("body").evaluate(() => {
        const elm = document.querySelector("div.dtable, #table-execution-task");
        if (!elm) return [];
        const dtableInstance = (window as any).zui?.DTable?.query(elm);
        if (!dtableInstance) return [];
        const rows = dtableInstance.options?.data || dtableInstance.data || [];
        return rows
          .map((r: any) => ({
            id: String(r.rawID || r.id || ""),
            title: String(r.name || r.title || ""),
            status: String(r.status || r.data?.status || r.taskStatus || ""),
          }))
          .filter((t: any) => t.id);
      });

      console.log(
        `[INFO] Found ${iterationTickets.length} tickets in iteration: ${iteration.name}`,
      );

      for (const ticket of iterationTickets) {
        if (processedTickets.has(ticket.id)) continue;
        processedTickets.add(ticket.id);

        const ticketUrl = `${BASE_URL}/zentao/task-view-${ticket.id}.html`;

        try {
          await navigateTo(page, ticketUrl);

          // ── FAST description extraction ──────────────────────────────
          const articleHtml = await extractDescriptionFast(page);

          let hasDescription = validateDescription(articleHtml);
          let status:
            | "Has Description"
            | "Missing Description"
            | "Has Description (Auto Edited)" = hasDescription
            ? "Has Description"
            : "Missing Description";

          if (
            !hasDescription &&
            process.env.AUTO_EDIT === "true" &&
            process.env.GEMINI_API_KEY
          ) {
            console.log(
              `[INFO] Ticket ${ticket.id} description is missing. Attempting auto-edit...`,
            );
            const generatedDesc = await generateDescriptionWithGemini(
              ticket.title,
            );
            if (generatedDesc) {
              const success = await autoEditTaskDescription(
                page,
                ticket.id,
                generatedDesc,
              );
              if (success) {
                status = "Has Description (Auto Edited)";
                hasDescription = true;
              }
            }
          }

          console.log(`[INFO] Ticket : ${ticket.id} | ${status}`);

          allTickets.push({
            id: ticket.id,
            iteration: iteration.name,
            title: ticket.title,
            taskStatus: mapStatus(ticket.status),
            status,
            url: ticketUrl,
          });

          // Polite crawl delay — reduced from 300ms
          await page.waitForTimeout(100);
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
            status: "Missing Description",
            url: ticketUrl,
          });
        }
      }
    }

    // ── 5. Generate Excel report ───────────────────────────────────────────
    const totalIterationsCount = iterationsToProcess.length;
    const totalTicketsCount = allTickets.length;
    const hasDescCount = allTickets.filter(
      (t) =>
        t.status === "Has Description" ||
        t.status === "Has Description (Auto Edited)",
    ).length;
    const missingDescCount = allTickets.filter(
      (t) => t.status === "Missing Description",
    ).length;
    const completionRate =
      totalTicketsCount > 0 ? (hasDescCount / totalTicketsCount) * 100 : 0;
    const missingRate =
      totalTicketsCount > 0 ? (missingDescCount / totalTicketsCount) * 100 : 0;

    console.log("\n[SUMMARY]");
    console.log(`Total Iteration     : ${totalIterationsCount}`);
    console.log(`Total Ticket        : ${totalTicketsCount}`);
    console.log(`Has Description     : ${hasDescCount}`);
    console.log(`Missing Description : ${missingDescCount}`);

    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Summary
    const summarySheet = workbook.addWorksheet("Summary");
    summarySheet.getColumn("A").width = 30;
    summarySheet.getColumn("B").width = 20;
    summarySheet.addRow(["Metric", "Value"]);
    summarySheet.addRow(["Total Iteration", totalIterationsCount]);
    summarySheet.addRow(["Total Ticket", totalTicketsCount]);
    summarySheet.addRow(["Ticket Dengan Description", hasDescCount]);
    summarySheet.addRow(["Ticket Tanpa Description", missingDescCount]);
    summarySheet.addRow(["Completion Rate", `${completionRate.toFixed(2)}%`]);
    summarySheet.addRow(["Missing Rate", `${missingRate.toFixed(2)}%`]);

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
      fgColor: { argb: "FF1F497D" },
    };
    summaryHeaderRow.alignment = { vertical: "middle", horizontal: "center" };

    for (let r = 2; r <= 7; r++) {
      const row = summarySheet.getRow(r);
      row.getCell("A").font = { bold: true };
      row.getCell("A").alignment = { horizontal: "left" };
      row.getCell("B").alignment = { horizontal: "right" };
      if (r === 6)
        row.getCell("B").font = { color: { argb: "FF008000" }, bold: true };
      else if (r === 7)
        row.getCell("B").font = { color: { argb: "FFFF0000" }, bold: true };
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
    for (const ticket of allTickets) allTicketsSheet.addRow(ticket);

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

    allTicketsSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const statusCell = row.getCell("E");
      const cellValue = String(statusCell.value || "");
      if (cellValue.startsWith("Has Description"))
        statusCell.font = { color: { argb: "FF008000" }, bold: true };
      else statusCell.font = { color: { argb: "FFFF0000" }, bold: true };
      row.getCell("D").alignment = { horizontal: "center" };
      if (rowNumber % 2 === 0) {
        row.eachCell((cell, colNumber) => {
          if (colNumber !== 5)
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFF2F5F8" },
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

    // Sheet 3: Missing Description
    const missingDescSheet = workbook.addWorksheet("Missing Description");
    missingDescSheet.columns = [
      { header: "Ticket ID", key: "id", width: 15 },
      { header: "Iteration", key: "iteration", width: 30 },
      { header: "Title", key: "title", width: 50 },
      { header: "Task Status", key: "taskStatus", width: 15 },
      { header: "URL", key: "url", width: 70 },
    ];
    const missingTickets = allTickets.filter(
      (t) => t.status === "Missing Description",
    );
    for (const ticket of missingTickets) missingDescSheet.addRow(ticket);

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
      fgColor: { argb: "FFC00000" },
    };
    missingHeaderRow.alignment = { vertical: "middle", horizontal: "center" };

    missingDescSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      row.getCell("D").alignment = { horizontal: "center" };
      if (rowNumber % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFF0F0" },
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
      console.log(`\n[INFO] Excel report generated: ${outputPath}`);
    } catch (writeError: any) {
      if (writeError.code === "EBUSY") {
        const hhmmss =
          String(date.getHours()).padStart(2, "0") +
          String(date.getMinutes()).padStart(2, "0") +
          String(date.getSeconds()).padStart(2, "0");
        reportFilename = `zentao_description_audit_${yyyy}${mm}${dd}_${hhmmss}.xlsx`;
        outputPath = path.join(process.cwd(), reportFilename);
        await workbook.xlsx.writeFile(outputPath);
        console.log(
          `\n[INFO] Excel report generated (fallback): ${outputPath}`,
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

runAudit();
