/**
 * Prompt template settings and the in-browser diagnostics.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, persistState, showToast, state, t } = App;

  /**
   * In-Browser Unit Diagnostics Runner
   */
  function initDiagnostics() {
    elements.diagnosticsBtn.addEventListener('click', function () {
      elements.diagnosticsModal.classList.add('open');
      runBrowserTests();
    });

    elements.closeDiagnosticsBtn.addEventListener('click', function () {
      elements.diagnosticsModal.classList.remove('open');
    });

    elements.runDiagnosticsBtn.addEventListener('click', function () {
      runBrowserTests();
    });
  }

  function runBrowserTests() {
    elements.diagnosticsLog.innerHTML = `<div class="text-muted">${I18n.getLanguage() === 'en' ? 'Running diagnostic suite...' : 'Spouštím interní diagnostickou sadu...'}</div>`;

    setTimeout(function () {
      const logs = [];
      let pass = 0;
      let fail = 0;

      function assertEq(name, actual, expected) {
        if (actual === expected) {
          logs.push(`<div class="diag-pass">✓ ${ComparatorEngine.escapeHtml(name)}</div>`);
          pass++;
        } else {
          logs.push(`<div class="diag-fail">✗ ${ComparatorEngine.escapeHtml(name)}<br><small class="diag-detail">Expected: ${ComparatorEngine.escapeHtml(String(expected))}, Got: ${ComparatorEngine.escapeHtml(String(actual))}</small></div>`);
          fail++;
        }
      }

      function assertOk(name, cond, details) {
        if (cond) {
          logs.push(`<div class="diag-pass">✓ ${ComparatorEngine.escapeHtml(name)}</div>`);
          pass++;
        } else {
          logs.push(`<div class="diag-fail">✗ ${ComparatorEngine.escapeHtml(name)} ${details ? '<br><small class="diag-detail">' + ComparatorEngine.escapeHtml(String(details)) + '</small>' : ''}</div>`);
          fail++;
        }
      }

      // 1. Classification
      assertEq('classifyRow(empty, empty) -> EMPTY', ComparatorEngine.classifyRow('', ''), 'EMPTY');
      assertEq('classifyRow(empty, text) -> ADDED', ComparatorEngine.classifyRow('', 'Nový text'), 'ADDED');
      assertEq('classifyRow(text, empty) -> REMOVED', ComparatorEngine.classifyRow('Starý text', ''), 'REMOVED');
      assertEq('classifyRow(text, text) -> UNCHANGED', ComparatorEngine.classifyRow('Etika 2024', 'Etika 2024'), 'UNCHANGED');
      assertEq('classifyRow(diff, diff) -> MODIFIED', ComparatorEngine.classifyRow('14 dnů', '7 dnů'), 'MODIFIED');

      // 2. Diffing
      const diffRes = ComparatorEngine.computeDiff('Lhůta 14 dnů.', 'Lhůta 7 dnů.');
      assertOk('computeDiff detects changes', diffRes.stats.hasChanges);
      assertOk('computeDiff contains removed token 14', diffRes.summary.removedTerms.includes('14'));
      assertOk('computeDiff contains added token 7', diffRes.summary.addedTerms.includes('7'));

      // 3. HTML Escaping
      const safeDiff = ComparatorEngine.renderDiffHtml(ComparatorEngine.computeDiff('<b>test</b>', '<i>test</i>'), 'inline');
      assertOk('renderDiffHtml sanitizes HTML tags', !safeDiff.includes('<script>') && safeDiff.includes('&lt;'));

      // 4. Stubs
      const sampleRows = SampleGenerator.getSampleData();
      const stubs = ComparatorEngine.generateTargetStub(sampleRows, 'CS 2024', 'DE 2024 (Referenz)', 'DE 2026 (Referenz)');
      assertEq('Stub generation for 15 sample paragraphs', stubs.length, 15);
      assertEq('Unchanged paragraph 1 marked auto_copied', stubs[0].status, 'auto_copied');

      // 5. Prompt Builder (Bilingual test)
      const promptEn = ComparatorEngine.buildCopilotPrompt({
        promptLang: 'en',
        targetLangName: 'English',
        classification: 'MODIFIED',
        refOld: 'DE Alt',
        refNew: 'DE Neu',
        targetOld: 'EN Old',
        diffResult: diffRes
      });
      assertOk('buildCopilotPrompt generates English prompt', promptEn.includes('professional translator') && promptEn.includes('CRITICAL INSTRUCTIONS'));

      // 6. Export prep
      const exportData = ComparatorEngine.prepareExportData({
        rows: sampleRows,
        headers: SampleGenerator.getSampleHeaders(),
        mappings: SampleGenerator.getSampleDefaultMappings(),
        translations: { 'CS 2024': stubs },
        exportType: 'clean'
      });
      assertOk('clean export leaves out removed paragraphs', exportData.data.length < sampleRows.length);

      const summaryHtml = `
        <div class="diag-summary">
          ${I18n.getLanguage() === 'en' ? 'Diagnostics Result:' : 'Výsledek diagnostiky:'} <span class="diag-pass">${pass} ${I18n.getLanguage() === 'en' ? 'passed' : 'úspěšných'}</span>, <span class="${fail > 0 ? 'diag-fail' : 'diag-pass'}">${fail} ${I18n.getLanguage() === 'en' ? 'failed' : 'chyb'}</span>
        </div>
      `;

      elements.diagnosticsLog.innerHTML = summaryHtml + logs.join('');
    }, 50);
  }

  /**
   * Prompt Template Settings Modal
   */
  function initPromptSettings() {
    elements.settingsPromptBtn.addEventListener('click', function () {
      elements.promptSettingsModal.classList.add('open');
      elements.promptTemplateInput.value = state.customPromptTemplate || getDefaultPromptSample();
    });

    elements.closePromptSettingsBtn.addEventListener('click', function () {
      elements.promptSettingsModal.classList.remove('open');
    });

    elements.savePromptTemplateBtn.addEventListener('click', function () {
      const value = elements.promptTemplateInput.value.trim();
      // Saving the untouched sample keeps the built-in (classification-specific) prompts
      state.customPromptTemplate = value && value !== getDefaultPromptSample().trim() ? value : null;
      persistState(0);
      elements.promptSettingsModal.classList.remove('open');
      showToast(t('promptSaveBtn'), 'success');
    });

    elements.resetPromptTemplateBtn.addEventListener('click', function () {
      state.customPromptTemplate = null;
      elements.promptTemplateInput.value = getDefaultPromptSample();
      persistState(0);
      showToast(t('promptResetBtn'), 'info');
    });
  }

  function getDefaultPromptSample() {
    if (I18n.getLanguage() === 'en') {
      return 'You are an experienced professional translator.\n\nTASK: Update the approved translation in language {TARGET_LANG} to match the new version of the reference text.\nParagraph: {ROW_ID}\n\n1. REFERENCE – OLD VERSION ({REF_LANG}):\n{REF_OLD}\n\n2. REFERENCE – NEW VERSION:\n{REF_NEW}\n{CHANGES_SUMMARY}\n\n3. CURRENT APPROVED TRANSLATION ({TARGET_LANG}):\n{TARGET_OLD}\n\nINSTRUCTIONS:\nMake ONLY the minimal changes reflecting the reference difference. Keep the existing tone and terminology. Output only the final text.';
    }
    return 'Jste zkušený odborný překladatel.\n\nÚKOL: Aktualizujte schválený překlad v jazyce {TARGET_LANG} podle nové verze referenčního textu.\nOdstavec: {ROW_ID}\n\n1. REFERENCE – STARÁ VERZE ({REF_LANG}):\n{REF_OLD}\n\n2. REFERENCE – NOVÁ VERZE:\n{REF_NEW}\n{CHANGES_SUMMARY}\n\n3. DOSAVADNÍ SCHVÁLENÝ PŘEKLAD ({TARGET_LANG}):\n{TARGET_OLD}\n\nPOKYNY:\nProveďte POUZE minimální změny odpovídající rozdílu v referenci. Zachovejte zavedenou terminologii a styl. Vraťte výhradně výsledný text.';
  }

  Object.assign(App, { initDiagnostics: initDiagnostics, initPromptSettings: initPromptSettings });
}(window.App = window.App || {}));
