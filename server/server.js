/**
 * AI 写作助手 - 本地 HTTPS 服务器
 * 打包成 server.exe 后，负责为 Word 插件提供 web 文件服务
 */

const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { execSync } = require('child_process');

// pkg 打包时，process.execPath 是 .exe 文件的真实路径
// __dirname 在 pkg 中指向虚拟文件系统，不能用于读取外部文件
const exeDir = path.dirname(process.execPath);
const WEB_DIR = path.join(exeDir, 'web');
const CERTS_DIR = path.join(exeDir, 'certs');
const PORT = 3000;
const LOG_FILE = path.join(exeDir, 'server.log');

// ---- 日志 ----
function log(msg) {
  const line = `[${new Date().toISOString()}] ${msg}\n`;
  process.stdout.write(line);
  try {
    fs.appendFileSync(LOG_FILE, line);
  } catch (e) { /* ignore */ }
}

// ---- MIME 类型 ----
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.xml': 'application/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

// ---- 请求处理 ----
function handleRequest(req, res) {
  // CORS 允许 Office 跨域访问
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('X-Content-Type-Options', 'nosniff');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // 解析请求路径
  const parsedUrl = url.parse(req.url);
  let filePath = path.join(WEB_DIR, parsedUrl.pathname);

  // 防止目录遍历攻击
  if (!filePath.startsWith(WEB_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  // 默认页面
  if (parsedUrl.pathname === '/' || parsedUrl.pathname === '') {
    filePath = path.join(WEB_DIR, 'src', 'taskpane', 'taskpane.html');
  }

  // 读取文件
  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) {
      // 尝试加 .html 扩展名
      const htmlPath = filePath + '.html';
      if (fs.existsSync(htmlPath)) {
        serveFile(htmlPath, res);
      } else {
        log(`404: ${req.url}`);
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Not Found: ' + req.url);
      }
      return;
    }
    serveFile(filePath, res);
  });
}

function serveFile(filePath, res) {
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      log(`Error reading file: ${filePath} - ${err.message}`);
      res.writeHead(500);
      res.end('Internal Server Error');
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}

// ---- 启动服务器 ----
function startServer() {
  const pfxFile = path.join(CERTS_DIR, 'server.pfx');
  const certFile = path.join(CERTS_DIR, 'server.crt');
  const keyFile = path.join(CERTS_DIR, 'server.key');
  const passphrase = 'WordAIAddin2024';

  let options = null;

  if (fs.existsSync(pfxFile)) {
    log('正在使用 PFX 格式证书启动 HTTPS 服务器...');
    try {
      options = {
        pfx: fs.readFileSync(pfxFile),
        passphrase: passphrase
      };
    } catch (err) {
      log(`加载 PFX 证书失败: ${err.message}`);
    }
  } else if (fs.existsSync(certFile) && fs.existsSync(keyFile)) {
    log('正在使用 CRT/KEY 格式证书启动 HTTPS 服务器...');
    try {
      options = {
        key: fs.readFileSync(keyFile),
        cert: fs.readFileSync(certFile),
      };
    } catch (err) {
      log(`加载 CRT/KEY 证书失败: ${err.message}`);
    }
  }

  if (!options) {
    log('ERROR: 找不到可用的 HTTPS 证书，回退到 HTTP 模式');
    startHttpServer();
    return;
  }

  // 检查 web 目录
  if (!fs.existsSync(WEB_DIR)) {
    log(`ERROR: web 目录不存在: ${WEB_DIR}`);
    process.exit(1);
  }

  try {
    const server = https.createServer(options, handleRequest);

    server.on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        log(`端口 ${PORT} 已被占用，服务器可能已在运行`);
        process.exit(0);
      }
      log(`服务器错误: ${err.message}`);
      process.exit(1);
    });

    server.listen(PORT, '127.0.0.1', () => {
      log(`AI 写作助手服务器启动成功`);
      log(`地址: https://localhost:${PORT}`);
      log(`Web目录: ${WEB_DIR}`);
    });
  } catch (err) {
    log(`启动 HTTPS 服务器失败: ${err.message}`);
    startHttpServer();
  }
}

function startHttpServer() {
  log('使用 HTTP 模式（调试用）...');
  const server = http.createServer(handleRequest);
  server.listen(PORT, '127.0.0.1', () => {
    log(`HTTP 服务器启动: http://localhost:${PORT}`);
  });
}

// ---- 优雅退出 ----
process.on('SIGTERM', () => {
  log('服务器正在关闭...');
  process.exit(0);
});

process.on('uncaughtException', (err) => {
  log(`未捕获异常: ${err.message}`);
  log(err.stack);
});

// 启动！
log('='.repeat(50));
log('AI 写作助手本地服务器 v1.0.0');
log('='.repeat(50));
startServer();
