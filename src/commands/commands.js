/* global Office, Word */
'use strict';

// =============================================
// 功能区命令处理器
// 处理来自 Word 功能区按钮的直接命令
// =============================================

Office.onReady(() => {
  // 注册所有命令函数
});

if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate('triggerAIInline', triggerAIInline);
}

/**
 * 触发 AI 行内补全（Ctrl+J 快捷键绑定）
 * @param {Office.AddinCommands.Event} event
 */
async function triggerAIInline(event) {
  try {
    const config = await loadConfig();
    if (!config.apiKey) {
      await showNotification('请先配置 API Key', '请点击 AI 菜单 → API 设置 配置您的 AI 服务');
      event.completed();
      return;
    }

    await Word.run(async context => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();

      const selectedText = selection.text;
      const prompt = '你是一位智能写作助手。请基于以下内容，生成一段自然流畅的补全文字，简洁到位，不超过100字。只输出补全内容。';
      const result = await callAI(config, prompt, selectedText || '请开始写作');

      selection.insertText(result, Word.InsertLocation.after);
      await context.sync();
    });
  } catch (err) {
    console.error('triggerAIInline error:', err);
  } finally {
    event.completed();
  }
}

// ---- 配置加载 ----
async function loadConfig() {
  try {
    const stored = localStorage.getItem('aiConfig');
    return stored ? JSON.parse(stored) : {};
  } catch {
    return {};
  }
}

// ---- AI 调用 ----
async function callAI(config, systemPrompt, userContent) {
  const baseUrl = (config.baseUrl || 'https://api.openai.com/v1').replace(/\/$/, '');
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`
    },
    body: JSON.stringify({
      model: config.model || 'gpt-4o',
      temperature: config.temperature || 0.7,
      max_tokens: config.maxTokens || 2048,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userContent }
      ]
    })
  });

  if (!response.ok) {
    throw new Error(`API Error: ${response.status}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content?.trim() || '';
}

// ---- 显示通知 ----
async function showNotification(title, message) {
  return Word.run(async context => {
    // 在状态栏显示通知（Word 不支持原生通知，记录到控制台）
    console.warn(`[AI 插件] ${title}: ${message}`);
    await context.sync();
  });
}

// 注册到全局（供 manifest 调用）
if (typeof global !== 'undefined') {
  global.triggerAIInline = triggerAIInline;
} else {
  window.triggerAIInline = triggerAIInline;
}
