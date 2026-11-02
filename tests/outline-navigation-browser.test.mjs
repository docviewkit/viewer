import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { chromium, webkit } from 'playwright-core';

const fixture = process.env.DOCVIEWKIT_OUTLINE_FIXTURE;

test('real DOCX outline navigates to the heading in continuous and single page modes', { skip: !fixture, timeout: 120_000 }, async () => {
  const server = process.env.DOCVIEWKIT_TEST_URL ? undefined : spawn(process.execPath, ['scripts/serve.mjs', '--port', '0'], { stdio: ['ignore', 'pipe', 'inherit'] });
  try {
    const url = process.env.DOCVIEWKIT_TEST_URL ?? await new Promise((resolve, reject) => {
      const lines = createInterface({ input: server.stdout });
      lines.on('line', line => { const match = line.match(/https?:\/\/[^\s]+/u); if (match) resolve(match[0]); });
      server.once('exit', code => reject(new Error(`Server exited: ${code}`)));
    });
    for (const engine of [chromium, webkit]) {
      const browser = await engine.launch({ headless: true });
      try {
        const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
        await page.goto(`${url.replace(/\/$/u, "")}/examples/viewer.html?fixture=visual-baseline.pptx`);
        await page.waitForFunction(() => document.documentElement.dataset.ready === 'true');
        await page.locator('#viewer-file-input').setInputFiles(fixture);
        await page.waitForFunction(() => document.querySelector('docviewkit-viewer').state.info?.format === 'docx');
        const viewer = page.locator('docviewkit-viewer');
        const items = await viewer.evaluate(v => v.state.info.outline);
        const targets = ['1.1 目的', '2.1 技术架构'];
        assert.equal(items.find(x => x.title === targets[0]).unitIndex, items.find(x => x.title === targets[1]).unitIndex);
        for (const mode of ['continuous', 'single']) {
          if (mode === 'single') await viewer.locator('.desktop-page-mode').click();
          for (const title of targets) {
            const item = items.find(x => x.title === title);
            assert.ok(Number.isFinite(item.y), `${title} must have a document position`);
            await viewer.locator('.outline-button').filter({ hasText: title }).click();
            await page.waitForFunction(({ title, mode }) => {
              const v = document.querySelector('docviewkit-viewer');
              const item = v.state.info.outline.find(x => x.title === title);
              const workspace = v.shadowRoot.querySelector('.workspace');
              const surface = mode === 'continuous' ? v.shadowRoot.querySelector(`.continuous-page[data-unit-index="${item.unitIndex}"]`) : v.shadowRoot.querySelector('.stage');
              if (!surface || !item || v.state.pageMode !== mode) return false;
              const unit = v.state.info.units[item.unitIndex];
              const top = surface.getBoundingClientRect().top - workspace.getBoundingClientRect().top + item.y * surface.getBoundingClientRect().width / unit.width;
              return Math.abs(top) < 3;
            }, { title, mode }, { timeout: 10000 });
          }
        }
        assert.deepEqual(errors, []);
      } finally { await browser.close(); }
    }
  } finally { server?.kill(); }
});
