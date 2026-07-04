/* global Office, Word */
'use strict';

// =============================================
// AI 写作助手 - 主逻辑
// =============================================

// ---- AI 配置管理 ----
const AIConfig = {
  /**
   * 从 Office 存储加载配置
   */
  load() {
    try {
      const stored = localStorage.getItem('aiConfig');
      return stored ? JSON.parse(stored) : this.getDefaults();
    } catch (e) {
      return this.getDefaults();
    }
  },

  getDefaults() {
    return {
      provider: 'openai',
      apiKey: '',
      baseUrl: 'https://api.openai.com/v1',
      model: 'gpt-4o',
      temperature: 0.7,
      maxTokens: 2048
    };
  }
};

// ---- AI 请求核心 ----
const AIClient = {
  /**
   * 调用 AI API
   * @param {string} systemPrompt - 系统提示词
   * @param {string} userContent - 用户内容
   * @returns {Promise<string>} AI 响应文本
   */
  async call(systemPrompt, userContent) {
    const config = AIConfig.load();

    if (!config.apiKey) {
      throw new Error('请先在 API 设置中配置您的 API Key');
    }

    const baseUrl = config.baseUrl.replace(/\/$/, '');
    const endpoint = `${baseUrl}/chat/completions`;

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${config.apiKey}`
      },
      body: JSON.stringify({
        model: config.model,
        temperature: config.temperature,
        max_tokens: config.maxTokens,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userContent }
        ]
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      const msg = errorData?.error?.message || `HTTP ${response.status}`;
      throw new Error(`API 请求失败：${msg}`);
    }

    const data = await response.json();
    return data.choices?.[0]?.message?.content?.trim() || '';
  }
};

// ---- AI 功能提示词 ----
const Prompts = {
  edit: {
    optimize: '你是一位专业的文字编辑。请优化以下文字的表达，使其更加清晰、流畅、有力，同时保持原意。只输出优化后的文字，不要解释。',
    'fix-grammar': '你是一位专业的语言编辑。请修复以下文字中的拼写错误和语法问题，保持原意和风格。只输出修复后的文字，不要解释。',
    shorten: '你是一位专业的文字编辑。请将以下文字精简压缩，去除冗余，保留核心信息，使表达更加简练。只输出精简后的文字，不要解释。',
    expand: '你是一位专业的文字编辑。请将以下文字扩展丰富，增加细节、例子和背景信息，使内容更加充实。只输出扩展后的文字，不要解释。',
    simplify: '你是一位专业的文字编辑。请将以下文字改写成更简单易懂的语言，避免专业术语，适合普通读者阅读。只输出改写后的文字，不要解释。'
  },
  tone: {
    professional: '你是一位专业的文字编辑。请将以下文字改写成正式、严谨的专业风格，措辞准确，逻辑清晰，适用于学术、科技、医疗、法律、政务、商务等各类专业领域。只输出改写后的文字，不要解释。',
    casual: '你是一位专业的文字编辑。请将以下文字改写成轻松、自然的口语风格，如同朋友聊天，亲切易读。只输出改写后的文字，不要解释。'
  },
  generate: {
    explain: '你是一位知识渊博的助手。请详细解释以下内容的含义、背景和重要性，语言清晰易懂。',
    summarize: '你是一位专业的文字编辑。请对以下内容进行摘要总结，提取核心要点，用简洁的语言呈现。',
    continue: '你是一位专业的写作助手。请根据以下内容的风格和主题，自然地续写后续内容，保持行文一致。只输出续写内容，不要重复原文。'
  }
};

// ---- 状态管理 ----
let currentResult = '';
let currentAction = '';

// ---- UI 工具函数 ----
const UI = {
  showPanel(panelId) {
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    const target = document.getElementById(`panel-${panelId}`);
    if (target) target.classList.add('active');
  },

  showLoading(text = 'AI 处理中...') {
    document.getElementById('loadingText').textContent = text;
    document.getElementById('loadingOverlay').classList.add('active');
  },

  hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
  },

  showResult(resultAreaId, text) {
    const area = document.getElementById(resultAreaId);
    if (area) {
      area.textContent = text;
      area.classList.add('visible');
    }
    document.getElementById('resultFooter').style.display = 'flex';
    currentResult = text;
  },

  hideResult(resultAreaId) {
    const area = document.getElementById(resultAreaId);
    if (area) {
      area.textContent = '';
      area.classList.remove('visible');
    }
    document.getElementById('resultFooter').style.display = 'none';
    currentResult = '';
  },

  toast(message, type = '') {
    const el = document.getElementById('toast');
    el.textContent = message;
    el.className = `toast show ${type}`;
    setTimeout(() => {
      el.className = 'toast';
    }, 3000);
  }
};

// ---- Word 文档操作 ----
const WordHelper = {
  /**
   * 获取当前选中的文字
   */
  async getSelectedText() {
    return new Promise((resolve, reject) => {
      Word.run(async context => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        resolve(selection.text);
      }).catch(reject);
    });
  },

  /**
   * 插入文字到光标位置
   */
  async insertText(text) {
    return Word.run(async context => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.after);
      await context.sync();
    });
  },

  /**
   * 替换选中文字
   */
  async replaceSelection(text) {
    return Word.run(async context => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  }
};

// ---- 核心 AI 操作处理器 ----
async function executeAIAction(systemPrompt, loadingText, resultAreaId) {
  let selectedText = '';

  try {
    selectedText = await WordHelper.getSelectedText();
  } catch (e) {
    // 如果无法获取选中文字，继续（用于续写等场景）
  }

  if (!selectedText && !resultAreaId.includes('continue')) {
    UI.toast('请先在文档中选中要处理的文字', 'error');
    return;
  }

  UI.showLoading(loadingText);

  try {
    const result = await AIClient.call(systemPrompt, selectedText || '（请继续写作）');
    UI.hideLoading();
    UI.showResult(resultAreaId, result);
  } catch (err) {
    UI.hideLoading();
    UI.toast(err.message || 'AI 请求失败，请检查设置', 'error');

    // 如果 API Key 未设置，自动打开设置页面
    if (err.message.includes('API Key')) {
      setTimeout(() => openSettings(), 1000);
    }
  }
}

// ---- 触发行内补全 ----
async function triggerInlineCompletion() {
  UI.showLoading('生成行内补全...');
  try {
    const selectedText = await WordHelper.getSelectedText();
    const prompt = '你是一位智能写作助手。请基于以下内容，生成一段自然流畅的补全文字，简洁到位，不超过100字。只输出补全内容。';
    const result = await AIClient.call(prompt, selectedText || '请开始写作');
    await WordHelper.insertText(result);
    UI.hideLoading();
    UI.toast('已插入 AI 补全内容 ✓', 'success');
  } catch (err) {
    UI.hideLoading();
    UI.toast(err.message, 'error');
    if (err.message.includes('API Key')) {
      setTimeout(() => openSettings(), 1000);
    }
  }
}

// ---- 打开设置 ----
function openSettings() {
  Office.context.ui.displayDialogAsync(
    `${window.location.href.split('/src/')[0]}/src/settings/settings.html`,
    { height: 70, width: 50 },
    result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // 设置页面已通过对话框打开
      }
    }
  );
}

// ---- 初始化 ----
Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    initUI();
    handleURLParams();
  }
});

function initUI() {
  // 导航按钮（功能卡片）
  document.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      if (action === 'trigger-inline') {
        triggerInlineCompletion();
      } else {
        UI.showPanel(action);
      }
    });
  });

  // 返回按钮
  document.querySelectorAll('.back-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      UI.showPanel(btn.dataset.target || 'home');
      // 清除结果
      document.querySelectorAll('.result-area').forEach(a => {
        a.textContent = '';
        a.classList.remove('visible');
      });
      document.getElementById('resultFooter').style.display = 'none';
      currentResult = '';
    });
  });

  // 设置按钮
  document.getElementById('openSettingsBtn').addEventListener('click', openSettings);

  // ---- 编辑子功能 ----
  document.querySelectorAll('[data-edit-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.editAction;
      executeAIAction(Prompts.edit[action], '优化写作中...', 'editResult');
    });
  });

  // ---- 语气子功能 ----
  document.querySelectorAll('[data-tone-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.toneAction;
      executeAIAction(Prompts.tone[action], '调整语气中...', 'toneResult');
    });
  });

  // ---- 语言选择 ----
  document.querySelectorAll('.lang-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.lang-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // ---- 翻译 ----
  document.getElementById('translateBtn').addEventListener('click', () => {
    const activeLang = document.querySelector('.lang-btn.active');
    const langMap = {
      zh: '中文', en: '英语', ja: '日语', ko: '韩语',
      fr: '法语', de: '德语', es: '西班牙语', ru: '俄语'
    };
    const targetLang = langMap[activeLang?.dataset.lang] || '中文';
    const prompt = `你是一位专业翻译。请将以下文字翻译成${targetLang}，保持原文的语气和风格。只输出翻译结果，不要解释。`;
    executeAIAction(prompt, `翻译为${targetLang}中...`, 'translateResult');
  });

  // ---- 生成子功能 ----
  document.querySelectorAll('[data-generate-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.generateAction;
      const loadingTexts = {
        explain: '分析解释中...',
        summarize: '生成摘要中...',
        continue: '续写内容中...'
      };
      executeAIAction(Prompts.generate[action], loadingTexts[action], 'generateResult');
    });
  });

  // ---- 自定义改写 ----
  document.getElementById('customRewriteBtn').addEventListener('click', () => {
    const instruction = document.getElementById('customInstruction').value.trim();
    if (!instruction) {
      UI.toast('请输入改写指令', 'error');
      return;
    }
    const prompt = `你是一位专业的文字编辑。请按照以下指令改写内容：${instruction}\n\n只输出改写后的文字，不要解释。`;
    executeAIAction(prompt, '自定义改写中...', 'customRewriteResult');
  });

  // ---- 底部操作：插入 ----
  document.getElementById('insertBtn').addEventListener('click', async () => {
    if (!currentResult) return;
    try {
      await WordHelper.insertText('\n' + currentResult);
      UI.toast('已插入到文档 ✓', 'success');
    } catch (e) {
      UI.toast('插入失败：' + e.message, 'error');
    }
  });

  // ---- 底部操作：替换 ----
  document.getElementById('replaceBtn').addEventListener('click', async () => {
    if (!currentResult) return;
    try {
      await WordHelper.replaceSelection(currentResult);
      UI.toast('已替换选中文字 ✓', 'success');
    } catch (e) {
      UI.toast('替换失败：' + e.message, 'error');
    }
  });

  // ---- 底部操作：复制 ----
  document.getElementById('copyBtn').addEventListener('click', () => {
    if (!currentResult) return;
    navigator.clipboard.writeText(currentResult).then(() => {
      UI.toast('已复制到剪贴板 ✓', 'success');
    }).catch(() => {
      // 降级方案
      const textarea = document.createElement('textarea');
      textarea.value = currentResult;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      UI.toast('已复制到剪贴板 ✓', 'success');
    });
  });
}

// 处理 URL 参数（从 manifest 中不同入口点打开）
function handleURLParams() {
  const params = new URLSearchParams(window.location.search);
  const action = params.get('action');
  if (action) {
    const panelMap = {
      'edit': 'edit',
      'tone': 'tone',
      'translate': 'translate',
      'generate': 'generate',
      'custom-rewrite': 'custom-rewrite'
    };
    if (panelMap[action]) {
      UI.showPanel(panelMap[action]);
    }
  }
}
