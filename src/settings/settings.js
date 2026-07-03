/* global Office */
'use strict';

// =============================================
// API 设置页面逻辑
// =============================================

const DEFAULT_BASE_URLS = {
  openai: 'https://api.openai.com/v1',
  azure: 'https://YOUR_RESOURCE_NAME.openai.azure.com',
  custom: ''
};

// ---- Toast 通知 ----
function showToast(message, type = '') {
  const el = document.getElementById('toast');
  el.textContent = message;
  el.className = `toast show ${type}`;
  setTimeout(() => { el.className = 'toast'; }, 3000);
}

// ---- 加载已保存配置 ----
function loadConfig() {
  try {
    const stored = localStorage.getItem('aiConfig');
    return stored ? JSON.parse(stored) : getDefaults();
  } catch {
    return getDefaults();
  }
}

function getDefaults() {
  return {
    provider: 'openai',
    apiKey: '',
    baseUrl: DEFAULT_BASE_URLS.openai,
    deployment: '',
    model: 'gpt-4o',
    customModel: '',
    temperature: 0.7,
    maxTokens: 2048
  };
}

// ---- 保存配置 ----
function saveConfig(config) {
  return new Promise((resolve) => {
    localStorage.setItem('aiConfig', JSON.stringify(config));
    // 同步保存到 document.settings 以防有些旧环境 localStorage 不稳定
    try {
      Office.context.document.settings.set('aiConfig', JSON.stringify(config));
      Office.context.document.settings.saveAsync();
    } catch(e) {}
    resolve();
  });
}

// ---- 收集当前表单值 ----
function collectConfig() {
  const provider = document.querySelector('input[name="provider"]:checked')?.value || 'openai';
  const modelSelect = document.getElementById('modelSelect').value;
  const model = modelSelect === 'custom'
    ? document.getElementById('customModel').value.trim()
    : modelSelect;

  return {
    provider,
    apiKey: document.getElementById('apiKey').value.trim(),
    baseUrl: document.getElementById('baseUrl').value.trim() || DEFAULT_BASE_URLS[provider],
    deployment: document.getElementById('deployment').value.trim(),
    model: model || 'gpt-4o',
    customModel: modelSelect === 'custom' ? model : '',
    temperature: parseFloat(document.getElementById('temperature').value),
    maxTokens: parseInt(document.getElementById('maxTokens').value)
  };
}

// ---- 填充表单 ----
function populateForm(config) {
  // 提供商
  const providerRadio = document.getElementById(`provider-${config.provider}`);
  if (providerRadio) providerRadio.checked = true;
  updateProviderUI(config.provider);

  // API Key
  document.getElementById('apiKey').value = config.apiKey || '';

  // Base URL
  document.getElementById('baseUrl').value = config.baseUrl || DEFAULT_BASE_URLS[config.provider] || '';

  // Deployment (Azure)
  document.getElementById('deployment').value = config.deployment || '';

  // 模型选择
  const modelSelect = document.getElementById('modelSelect');
  const knownModels = ['gpt-4o', 'gpt-4o-mini', 'gpt-4-turbo', 'gpt-4', 'gpt-3.5-turbo'];
  if (config.customModel || !knownModels.includes(config.model)) {
    modelSelect.value = 'custom';
    document.getElementById('customModel').value = config.model || '';
    document.getElementById('customModelGroup').style.display = 'block';
  } else {
    modelSelect.value = config.model || 'gpt-4o';
  }

  // Temperature
  const tempEl = document.getElementById('temperature');
  tempEl.value = config.temperature ?? 0.7;
  document.getElementById('tempValue').textContent = tempEl.value;
  updateSliderFill(tempEl);

  // Max Tokens
  document.getElementById('maxTokens').value = config.maxTokens || 2048;
}

// ---- 根据提供商更新 UI ----
function updateProviderUI(provider) {
  const deploymentGroup = document.getElementById('deploymentGroup');
  const baseUrlEl = document.getElementById('baseUrl');

  deploymentGroup.style.display = provider === 'azure' ? 'block' : 'none';

  if (provider !== 'custom') {
    const currentBase = baseUrlEl.value;
    if (!currentBase || Object.values(DEFAULT_BASE_URLS).includes(currentBase)) {
      baseUrlEl.value = DEFAULT_BASE_URLS[provider];
    }
  }
}

// ---- 更新滑块填充色 ----
function updateSliderFill(slider) {
  const pct = (parseFloat(slider.value) / parseFloat(slider.max)) * 100;
  slider.style.background = `linear-gradient(to right, #6C63FF 0%, #6C63FF ${pct}%, #1E1E30 ${pct}%)`;
}

// ---- 测试连接 ----
async function testConnection() {
  const config = collectConfig();
  const resultEl = document.getElementById('testResult');
  const btnText = document.getElementById('testBtnText');

  if (!config.apiKey) {
    resultEl.className = 'test-result error';
    resultEl.textContent = '❌ 请先填写 API Key';
    return;
  }

  btnText.textContent = '⏳ 测试中...';
  document.getElementById('testConnectionBtn').disabled = true;
  resultEl.className = 'test-result';

  try {
    const baseUrl = (config.baseUrl || DEFAULT_BASE_URLS[config.provider]).replace(/\/$/, '');
    const endpoint = `${baseUrl}/models`;

    const response = await fetch(endpoint, {
      headers: {
        'Authorization': `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      const modelCount = data.data?.length || 0;
      resultEl.className = 'test-result success';
      resultEl.textContent = `✅ 连接成功！发现 ${modelCount} 个可用模型。`;
    } else if (response.status === 401) {
      resultEl.className = 'test-result error';
      resultEl.textContent = '❌ 认证失败：API Key 无效或已过期';
    } else if (response.status === 404) {
      // 尝试用 chat/completions 测试（某些自定义服务不支持 /models）
      await testWithChatAPI(config, resultEl);
    } else {
      resultEl.className = 'test-result error';
      resultEl.textContent = `❌ 连接失败：HTTP ${response.status}`;
    }
  } catch (err) {
    resultEl.className = 'test-result error';
    resultEl.textContent = `❌ 网络错误：${err.message}`;
  } finally {
    btnText.textContent = '🔗 测试连接';
    document.getElementById('testConnectionBtn').disabled = false;
  }
}

async function testWithChatAPI(config, resultEl) {
  try {
    const baseUrl = (config.baseUrl || DEFAULT_BASE_URLS[config.provider]).replace(/\/$/, '');
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: config.model,
        max_tokens: 5,
        messages: [{ role: 'user', content: 'hi' }]
      })
    });

    if (response.ok) {
      resultEl.className = 'test-result success';
      resultEl.textContent = '✅ 连接成功！API 响应正常。';
    } else if (response.status === 401) {
      resultEl.className = 'test-result error';
      resultEl.textContent = '❌ 认证失败：API Key 无效或已过期';
    } else {
      const errData = await response.json().catch(() => ({}));
      resultEl.className = 'test-result error';
      resultEl.textContent = `❌ 失败：${errData?.error?.message || `HTTP ${response.status}`}`;
    }
  } catch (err) {
    resultEl.className = 'test-result error';
    resultEl.textContent = `❌ 无法连接到服务器：${err.message}`;
  }
}

// ---- 初始化 ----
Office.onReady(() => {
  const config = loadConfig();
  populateForm(config);

  // 提供商切换
  document.querySelectorAll('input[name="provider"]').forEach(radio => {
    radio.addEventListener('change', () => {
      updateProviderUI(radio.value);
    });
  });

  // Temperature 滑块
  const tempSlider = document.getElementById('temperature');
  tempSlider.addEventListener('input', () => {
    document.getElementById('tempValue').textContent = tempSlider.value;
    updateSliderFill(tempSlider);
  });

  // 模型选择切换
  document.getElementById('modelSelect').addEventListener('change', function() {
    document.getElementById('customModelGroup').style.display =
      this.value === 'custom' ? 'block' : 'none';
  });

  // 显示/隐藏 API Key
  document.getElementById('toggleApiKey').addEventListener('click', () => {
    const input = document.getElementById('apiKey');
    const icon = document.getElementById('eyeIcon');
    if (input.type === 'password') {
      input.type = 'text';
      icon.innerHTML = `
        <path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94"/>
        <path d="M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19"/>
        <line x1="1" y1="1" x2="23" y2="23"/>
      `;
    } else {
      input.type = 'password';
      icon.innerHTML = `
        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
        <circle cx="12" cy="12" r="3"/>
      `;
    }
  });

  // 测试连接
  document.getElementById('testConnectionBtn').addEventListener('click', testConnection);

  // 取消
  document.getElementById('cancelBtn').addEventListener('click', () => {
    // 任务窗格无法用代码直接关闭，提示用户手动切换或关闭
    showToast('操作已取消，请点击其他功能按钮或关闭此面板', 'warning');
  });

  // 保存
  document.getElementById('saveBtn').addEventListener('click', async () => {
    const config = collectConfig();
    const saveText = document.getElementById('saveBtnText');

    if (!config.apiKey) {
      showToast('请填写 API Key', 'error');
      document.getElementById('apiKey').focus();
      return;
    }

    saveText.textContent = '保存中...';
    document.getElementById('saveBtn').disabled = true;

    try {
      await saveConfig(config);
      saveText.textContent = '✅ 已保存';
      showToast('设置保存成功！', 'success');

      setTimeout(() => {
        saveText.textContent = '💾 保存设置';
        document.getElementById('saveBtn').disabled = false;
      }, 1500);
    } catch (err) {
      saveText.textContent = '💾 保存设置';
      document.getElementById('saveBtn').disabled = false;
      showToast('保存失败：' + err.message, 'error');
    }
  });
});
