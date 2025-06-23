const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const Dotenv = require('dotenv-webpack');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const webpack = require('webpack');
const isDev = process.env.NODE_ENV !== 'production';

let apiUrl;

switch (process.env.NODE_ENV) {
  case 'production':
    apiUrl = 'https://flowizio.com/api';
    break;
  case 'qa':
    apiUrl = 'https://qa.flowizio.com/api';
    break;
  default:
    apiUrl = 'http://localhost:5000/api';
}

module.exports = {
  entry: './src/main.tsx',
  output: {
    path: path.resolve(__dirname, '../dist'),
    filename: isDev ? '[name].js' : '[name].[contenthash].js',
    publicPath: isDev ? '/' : '/Machine-Work-Log-Card/', // ✅ Виправлено
    clean: true,
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.js'],
    alias: {
      '@': path.resolve(__dirname, '../src'),
    },
  },
  devtool: isDev ? 'inline-source-map' : false,
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
      {
        test: /\.s[ac]ss$/i,
        include: path.resolve(__dirname, '../src'), // ⬅️ гарантує, що loader застосовується лише до твоїх файлів
        use: [
          isDev ? 'style-loader' : MiniCssExtractPlugin.loader,
          {
            loader: 'css-loader',
            options: {
              modules: {
                auto: (resPath) => resPath.includes('.module.'),
                localIdentName: isDev
                  ? '[path][name]__[local]--[hash:base64:5]'
                  : '[hash:base64:8]',
              },
            },
          },
          {
            loader: 'sass-loader',
            options: {
              // ✅ Виправляє warning про legacy JS API
              implementation: require('sass'),
              sourceMap: isDev,
            },
          },
        ],
      },
      {
        test: /\.(png|jpe?g|gif)$/i,
        type: 'asset/resource',
        generator: {
          filename: 'images/[name][hash][ext][query]',
        },
      },
      {
        test: /\.svg$/i,
        oneOf: [
          {
            issuer: /\.[jt]sx?$/,
            use: ['@svgr/webpack'],
          },
          {
            type: 'asset/resource',
            generator: {
              filename: 'images/[name][hash][ext][query]',
            },
          },
        ],
      },
      {
        test: /\.(woff(2)?|eot|ttf|otf)$/i,
        type: 'asset/resource',
        generator: {
          filename: 'fonts/[name][hash][ext][query]',
        },
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, '../public/index.html'),
      favicon: false,
    }),
    new webpack.DefinePlugin({
      __API__: JSON.stringify(apiUrl),
      __IS_DEV__: JSON.stringify(isDev),
    }),
    new Dotenv(),
    !isDev &&
      new MiniCssExtractPlugin({
        filename: 'styles/[name].[contenthash].css',
      }),
  ].filter(Boolean),
  devServer: {
    static: {
      directory: path.join(__dirname, '../public'),
    },
    compress: true,
    port: 3000,
    historyApiFallback: true,
    open: true,
  },
  mode: isDev ? 'development' : 'production',
};
