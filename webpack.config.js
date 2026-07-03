const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, options) => {
  const isDev = options.mode === 'development';

  return {
    devtool: 'source-map',
    entry: {
      taskpane: './src/taskpane/taskpane.js',
      settings: './src/settings/settings.js',
      commands: './src/commands/commands.js'
    },

    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },

    devServer: {
      static: {
        directory: path.join(__dirname, 'dist')
      },
      server: {
        type: 'https',
        options: {
          // 使用 office-addin-dev-certs 生成的证书
          ca: `${process.env.USERPROFILE}/.office-addin-dev-certs/ca.crt`,
          key: `${process.env.USERPROFILE}/.office-addin-dev-certs/localhost.key`,
          cert: `${process.env.USERPROFILE}/.office-addin-dev-certs/localhost.crt`
        }
      },
      port: 3000,
      headers: {
        'Access-Control-Allow-Origin': '*'
      }
    },

    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'src/taskpane/taskpane.html',
        chunks: ['taskpane'],
        inject: false  // 使用 HTML 中已有的 script 标签
      }),
      new HtmlWebpackPlugin({
        template: './src/settings/settings.html',
        filename: 'src/settings/settings.html',
        chunks: ['settings'],
        inject: false
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'src/commands/commands.html',
        chunks: ['commands'],
        inject: false
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'assets', to: 'assets' },
          { from: 'src/taskpane/taskpane.css', to: 'src/taskpane/taskpane.css' },
          { from: 'src/settings/settings.css', to: 'src/settings/settings.css' }
        ]
      })
    ],

    resolve: {
      extensions: ['.js']
    }
  };
};
