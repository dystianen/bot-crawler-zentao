import { chromium } from "playwright";

const BASE_URL = "http://pm.solusi-ku.id";
const USERNAME = "dystian.en";
const PASSWORD = "Solusiku123";
const TICKET_URL = `${BASE_URL}/zentao/task-view-920.html`;

async function runDiagnostic() {
  console.log("Starting diagnostic...");
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
  });
  const page = await context.newPage();

  try {
    console.log(`Navigating to initial page to log in...`);
    await page.goto(`${BASE_URL}/zentao/execution-task-34-all-0-status,id_desc-17-100.html`, { waitUntil: "load" });
    
    // Login if needed
    const accountInput = page.locator("#account");
    if (await accountInput.isVisible()) {
      await page.fill("#account", USERNAME);
      await page.fill("#password", PASSWORD);
      await page.locator("#submit").click({ force: true });
      await page.waitForTimeout(5000);
      console.log("Logged in successfully.");
    }

    console.log(`Navigating to ticket URL: ${TICKET_URL}`);
    await page.goto(TICKET_URL, { waitUntil: "load" });

    // Wait for the iframe's task description to be visible
    const iframe = page.frameLocator('#appIframe-execution');
    const descSection = iframe.locator('div[zui-key="Task Description"], .detail-section');
    
    console.log("Waiting for 'Task Description' inside the iframe...");
    await descSection.first().waitFor({ state: "visible", timeout: 15000 });
    console.log("Task Description element is now visible inside iframe.");

    const count = await descSection.count();
    console.log(`Count of detail sections inside iframe: ${count}`);

    if (count > 0) {
      const article = iframe.locator('.article').first();
      const text = await article.innerText();
      console.log(`Article text: "${text}"`);
      const html = await article.innerHTML();
      console.log(`Article HTML: "${html}"`);
    }

  } catch (err) {
    console.error("Diagnostic error:", err);
  } finally {
    await browser.close();
    console.log("Diagnostic complete.");
  }
}

runDiagnostic();
