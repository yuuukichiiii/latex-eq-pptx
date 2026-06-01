const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

module.exports = async (env, options) => {
  const isDev = options.mode !== 'production';

  let httpsOptions = true; // Use webpack-dev-server's built-in self-signed cert by default
  try {
    const devCerts = require('office-addin-dev-certs');
    httpsOptions = await devCerts.getHttpsServerOptions();
  } catch {
    // fall back to self-signed
  }

  return {
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
    },
    output: {
      filename: '[name].bundle.js',
      path: path.resolve(__dirname, 'dist'),
      clean: true,
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: 'ts-loader',
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [MiniCssExtractPlugin.loader, 'css-loader'],
        },
        {
          // KaTeX fonts
          test: /\.(woff|woff2|eot|ttf|otf)$/i,
          type: 'asset/resource',
          generator: {
            filename: 'fonts/[name][ext]',
          },
        },
      ],
    },
    resolve: {
      extensions: ['.tsx', '.ts', '.js'],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane'],
      }),
      new MiniCssExtractPlugin({
        filename: '[name].css',
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'assets', to: 'assets', noErrorOnMissing: true },
        ],
      }),
    ],
    devServer: {
      port: 3000,
      server: {
        type: 'https',
        options: httpsOptions,
      },
      headers: { 'Access-Control-Allow-Origin': '*' },
      hot: true,
    },
    devtool: isDev ? 'source-map' : false,
  };
};
