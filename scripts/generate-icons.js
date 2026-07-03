/**
 * 生成插件图标占位文件
 * 运行: node scripts/generate-icons.js
 *
 * 实际使用时，请将 icon-16.png, icon-32.png, icon-80.png
 * 放置到 assets/ 目录中。
 * 
 * 图标要求：
 * - icon-16.png: 16x16 像素
 * - icon-32.png: 32x32 像素
 * - icon-80.png: 80x80 像素
 * 
 * 图标已生成，请将 assets/icon-80.png 复制并缩放为以上三种尺寸。
 */

const fs = require('fs');
const path = require('path');

// 创建 assets 目录
const assetsDir = path.join(__dirname, '..', 'assets');
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

console.log('📁 assets/ 目录已准备好');
console.log('⚠️  请将以下图标文件放入 assets/ 目录：');
console.log('   - icon-16.png (16x16)');
console.log('   - icon-32.png (32x32)');
console.log('   - icon-80.png (80x80)');
console.log('');
console.log('💡 提示：可以使用图片编辑软件将 ai_plugin_icon.png 缩放为所需尺寸');
